Attribute VB_Name = "modVBStatic"
Option Explicit
'==============================================================================
' VBStatic  -  statically embed MSVBVM60.DLL into the OBJ files of a VB6 project
'
' Usage (command line):
'     VBStatic.exe [/DLL:ACTIVEX | /DLL:STANDARD] [/NOICON] [/OUT:<file>] <msvbvm60.dll> <main.OBJ> [chainSymbol]
'
'   <main.OBJ>   the project's main object (the one that holds the runtime imports,
'                e.g. out1.OBJ / Projekt1.OBJ). Class/module OBJs are linked unchanged.
'   /NOICON      keep the runtime's icons from becoming the application icon in Explorer.
'                The runtime's RT_ICON/RT_GROUP_ICON resources (IDs 1201-1204) are NOT removed:
'                msvbvm60 loads them itself with LoadImageA, and without them the app dies
'                with 'Out of memory'. Instead a harmless empty RT_GROUP_ICON (ID 1) is added
'                in front of them when the OBJ has no icon of its own, so Explorer's
'                "first icon" is that empty one and it falls back to the generic icon.
'   /OUT:<file>  write the patched object to <file>. Without /OUT the main OBJ is
'                overwritten in place (it is built completely in memory first; no temp files).
'   chainSymbol  the project's own entry point. ONE default (___vbaS) for EXE and DLL;
'                it is the symbol every VB6 main OBJ defines. Override only if needed.
'   /DLL:ACTIVEX   ActiveX DLL: keeps DllCanUnloadNow, DllGetClassObject,
'                  DllRegisterServer, DllUnregisterServer (the -export: directives
'                  of the main OBJ).
'   /DLL:STANDARD  plain DLL: removes every -export: directive of the main OBJ and turns
'                  the Dll* ActiveX entry functions into static symbols, so you can add
'                  exports yourself, e.g.  /EXPORT:B  (plain name works; use
'                  B=?B@Module1@@AAGXXZ only if the name is ambiguous)
'
' Output goes to a console (AllocConsole). Optional: run MakeConsole.cmd after compiling
' to switch VBStatic.exe to the console subsystem, then it prints into the calling
' cmd window (and works with redirection, e.g. VBStatic ... > log.txt).
'
' The result is ONE object: the main OBJ with the whole VB runtime fused into it.
' Link it with the project's other OBJs (no msvbvm60.lib, no default libs):
'     link /OUT:app.exe /SUBSYSTEM:WINDOWS /NODEFAULTLIB /ENTRY:_VBStaticEntry ^
'          out1.obj out2.obj ...
' DLL (/DLL:ACTIVEX or /DLL:STANDARD):
'     link /DLL /OUT:app.dll /NODEFAULTLIB /ENTRY:_VBStaticDllEntry@12 ^
'          out1.obj out2.obj ...
' A patched OBJ is detected (section .vbrt) and refused, so running the tool twice on the
' same file is harmless.
'
' How it works
'   1. The main OBJ is loaded into a writer (sections, relocations, symbols, strings),
'      then everything below is added to / changed in that one object.
'   2. The DLL's sections .text + ENGINE are merged into ONE section (.vbrt), because
'      the code contains rel32 jumps between them that have no base relocation.
'      .data becomes .vbdat.
'   3. Every HIGHLOW entry of the DLL's .reloc table becomes a COFF DIR32 relocation,
'      the stored value becomes "offset inside target section".
'   4. The DLL's own imports (KERNEL32, USER32, ...) are rebuilt as .idata$2/3/4/5/6
'      so the linker produces the import table for the final EXE.
'   5. The main OBJ's msvbvm60 import sections are emptied and renamed (.vbnulN); its IAT
'      (.idata$5) is turned into plain data (.vbiat) whose slots carry DIR32 relocations
'      straight into .vbrt. The main OBJ keeps defining __imp_* / thunk symbols for the
'      other OBJs.
'   6. For every other named export that the OBJ does not define, _Name / __imp__Name
'      are added (an existing undefined entry of the OBJ is filled in place).
'   7. __VBStaticEntry calls the runtime's DllMain(hInstance, DLL_PROCESS_ATTACH, 0)
'      and then jumps to the project's real entry (___vbaS).
'   8. The runtime's resources (.rsrc: string tables, typelibs, version ...) are merged
'      into the resource tree of the OBJ (.rsrc$01/.rsrc$02), or added as new sections
'      if the OBJ has none. Resources that already exist in the OBJ (same type/name/
'      language) win and are listed in the log. With /NOICON an empty RT_GROUP_ICON
'      (ID 1) is put in front of the runtime's icons when the OBJ has no icon of its own.
'   9. /DLL:xxx builds the entry for a DLL:  __VBStaticDllEntry@12  forwards
'      (hInst, reason, reserved) to the runtime's DllMain and to the project's own
'      DLL entry (___vbaS, which tail-jumps into the runtime's UserDllMain), in the
'      right order for attach / detach.
'  10. /DLL:STANDARD strips the ActiveX interface (see above).
'  11. The three globals of VBAEXE6.LIB(natsupp.obj) (__adjust_fdiv, __fltused, __except_list) are all
'      defined by the patched OBJ, so the linker never extracts natsupp.obj (no LNK2005 even when
'      VBAEXE6.LIB is on the link line or in the OBJ's default-lib directives).
'==============================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Sub ExitProcess Lib "Kernel32" (ByVal uExitCode As Long)
Private Declare Function AllocConsole Lib "Kernel32" () As Long
Private Declare Function FreeConsole Lib "Kernel32" () As Long
Private Declare Function GetStdHandle Lib "Kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function SetConsoleTitle Lib "Kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsole Lib "Kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Sub SetConsoleTextAttribute Lib "Kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long)
Private Declare Function SetConsoleMode Lib "Kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Const STD_OUTPUT_HANDLE As Long = -11

Private Const R_DIR32 As Long = 6
Private Const R_DIR32NB As Long = 7
Private Const R_REL32 As Long = &H14&

Private Type TBuf
    d() As Byte
    n As Long
    cap As Long
End Type

Private Type TRel
    va As Long
    sym As Long
    typ As Long
End Type

Private Type TSec
    nm As String
    flags As Long
    data As TBuf
    rel() As TRel
    nrel As Long
    relcap As Long
    symIdx As Long               '-1 = no section symbol known
    bss As Long                  'size of an uninitialised section (no raw data)
End Type

Private Type THash               'string -> Long hash table (replaces Scripting.Dictionary)
    cap As Long                  'power of two
    n As Long
    k() As String
    v() As Long
    u() As Byte                  '1 = slot in use
End Type

Private Type TRN                 'resource tree node
    isDir As Boolean
    isNamed As Boolean
    nameStr As String
    id As Long
    child As Long
    sib As Long
    nChild As Long
    dOff As Long                 'leaf: file offset of the data in its source array
    dSize As Long
    cp As Long
    srcDll As Boolean
    off As Long                  'offset inside the new .rsrc$01
End Type

'---- parsed DLL -------------------------------------------------------------
Private dll() As Byte
Private imgBase As Long, epRva As Long
Private nS As Long
Private sName(0 To 63) As String
Private sVA(0 To 63) As Long, sVS(0 To 63) As Long, sRP(0 To 63) As Long, sRS(0 To 63) As Long
Private ddRVA(0 To 15) As Long, ddSz(0 To 15) As Long
Private expBase As Long, expCnt As Long
Private expRva() As Long, expName() As String
Private hNames As THash                      'export name -> ordinal
Private nImp As Long
Private impDll(1 To 64) As String, impFT(1 To 64) As Long, impFirst(1 To 64) As Long, impCnt(1 To 64) As Long
Private nEnt As Long
Private entOrd(1 To 16384) As Long, entHint(1 To 16384) As Long, entName(1 To 16384) As String
Private relRva() As Long, nRel As Long

'---- parsed out1.OBJ ---------------------------------------------------------
Private ob() As Byte
Private oNs As Long, oSymp As Long, oNsym As Long, oStOff As Long, oStLen As Long
Private oNm() As String, oVal() As Long, oSec() As Long, oCls() As Long
Private oHave As THash, oDefined As THash

'---- resources ---------------------------------------------------------------
Private rn(1 To 4096) As TRN
Private rnN As Long
Private hLeaf As THash
Private resConflicts As String
Private resMode As Long                      '0 none, 1 merged into the OBJ's resource sections, 2 added as new sections
Private rsr01 As TBuf, rsr02 As TBuf, rsrRel As TBuf
Private rLeaf(1 To 4096) As Long, rLeafN As Long
Private hStr As THash, rCur As Long

'---- COFF writer state --------------------------------------------------------
Private cs() As TSec
Private csN As Long
Private symTab As TBuf, strTab As TBuf, symN As Long
Private codeLo As Long, codeHi As Long, dVA As Long, dVS As Long, iatRva As Long, iatSz As Long
Private sA As Long, sB As Long, sT As Long, sM As Long, sD2 As Long, sD3 As Long, sD4 As Long, sD5 As Long, sD6 As Long, sE As Long
Private mRc As Long                                 'process exit code: 0 = ok, 1 = error
Private hOut As Long, ownConsole As Boolean, outIsConsole As Boolean
Private gStrip As Boolean                    'True = /DLL:STANDARD (strip ActiveX exports)
Private gNoIcon As Boolean                   'True = /NOICON (hide the runtime's icons from Explorer)

'==============================================================================
' byte-buffer helpers
'==============================================================================
Private Sub BEnsure(b As TBuf, ByVal extra As Long)
    Dim need As Long, nc As Long
    need = b.n + extra
    If need <= b.cap Then Exit Sub
    nc = need * 2 + 256
    If b.cap = 0 Then
        ReDim b.d(0 To nc - 1)
    Else
        ReDim Preserve b.d(0 To nc - 1)
    End If
    b.cap = nc
End Sub

Private Sub BPutB(b As TBuf, ByVal v As Long)
    BEnsure b, 1
    b.d(b.n) = CByte(v And &HFF&)
    b.n = b.n + 1
End Sub

Private Sub BPutW(b As TBuf, ByVal v As Long)
    BEnsure b, 2
    b.d(b.n) = CByte(v And &HFF&)
    b.d(b.n + 1) = CByte((v And &HFF00&) \ &H100&)
    b.n = b.n + 2
End Sub

Private Sub BPutL(b As TBuf, ByVal v As Long)
    BEnsure b, 4
    CopyMemory b.d(b.n), v, 4
    b.n = b.n + 4
End Sub

Private Sub BPutBytes(b As TBuf, src() As Byte, ByVal off As Long, ByVal cnt As Long)
    If cnt <= 0 Then Exit Sub
    BEnsure b, cnt
    CopyMemory b.d(b.n), src(off), cnt
    b.n = b.n + cnt
End Sub

Private Sub BPutStr(b As TBuf, ByVal s As String)
    Dim i As Long
    For i = 1 To Len(s)
        BPutB b, Asc(Mid$(s, i, 1))
    Next
End Sub

Private Sub BZeros(b As TBuf, ByVal cnt As Long)
    If cnt <= 0 Then Exit Sub
    BEnsure b, cnt            'fresh memory from ReDim is zero
    b.n = b.n + cnt
End Sub

Private Sub BPad(b As TBuf, ByVal align As Long)
    Do While (b.n Mod align) <> 0
        BPutB b, 0
    Loop
End Sub

Private Sub SetName8(b As TBuf, ByVal off As Long, ByVal nm As String)
    Dim k As Long
    For k = 0 To 7
        If k < Len(nm) Then b.d(off + k) = Asc(Mid$(nm, k + 1, 1)) Else b.d(off + k) = 0
    Next
End Sub

Private Sub BPatchW(b As TBuf, ByVal off As Long, ByVal v As Long)
    b.d(off) = CByte(v And &HFF&)
    b.d(off + 1) = CByte((v And &HFF00&) \ &H100&)
End Sub

Private Sub BPutHex(b As TBuf, ByVal hexs As String)
    Dim i As Long, t As String
    hexs = Replace(hexs, " ", "")
    For i = 1 To Len(hexs) Step 2
        t = Mid$(hexs, i, 2)
        BPutB b, CLng("&H" & t)
    Next
End Sub

Private Sub BPatchL(b As TBuf, ByVal off As Long, ByVal v As Long)
    CopyMemory b.d(off), v, 4
End Sub

Private Function TW16(b As TBuf, ByVal off As Long) As Long
    TW16 = CLng(b.d(off)) Or (CLng(b.d(off + 1)) * &H100&)
End Function

Private Function BGetL(b As TBuf, ByVal off As Long) As Long
    Dim v As Long
    CopyMemory v, b.d(off), 4
    BGetL = v
End Function

Private Function L32(a() As Byte, ByVal off As Long) As Long
    Dim v As Long
    CopyMemory v, a(off), 4
    L32 = v
End Function

Private Function W16(a() As Byte, ByVal off As Long) As Long
    W16 = CLng(a(off)) Or (CLng(a(off + 1)) * &H100&)
End Function

Private Function CStrAt(a() As Byte, ByVal off As Long, Optional ByVal maxLen As Long = 4096) As String
    Dim i As Long, s As String
    i = off
    Do While i < off + maxLen
        If a(i) = 0 Then Exit Do
        s = s & Chr$(a(i))
        i = i + 1
    Loop
    CStrAt = s
End Function

Private Function ReadAll(ByVal path As String) As Byte()
    Dim f As Integer, b() As Byte
    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) = 0 Then Close #f: Err.Raise vbObjectError + 1, , "Empty file: " & path
    ReDim b(0 To LOF(f) - 1)
    Get #f, , b
    Close #f
    ReadAll = b
End Function

Private Sub WriteBuf(ByVal path As String, b As TBuf)
    Dim f As Integer, t() As Byte
    ReDim t(0 To b.n - 1)
    CopyMemory t(0), b.d(0), b.n
    If Len(Dir$(path)) > 0 Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, , t
    Close #f
End Sub

'==============================================================================
' tiny string -> Long hash table (open addressing, binary compare)
'==============================================================================
Private Function HashStr(ByVal s As String) As Long
    Dim i As Long, h As Long
    h = 5381
    For i = 1 To Len(s)
        h = ((h * 33) + (AscW(Mid$(s, i, 1)) And &HFFFF&)) And &H3FFFFF       'stays below 2^28: no overflow
    Next
    HashStr = h
End Function

Private Sub HInit(h As THash)
    h.cap = 256: h.n = 0
    ReDim h.k(0 To 255): ReDim h.v(0 To 255): ReDim h.u(0 To 255)
End Sub

Private Function HSlot(h As THash, ByVal key As String) As Long   'slot of key or -1
    Dim i As Long
    i = HashStr(key) And (h.cap - 1)
    Do While h.u(i) <> 0
        If h.k(i) = key Then HSlot = i: Exit Function
        i = (i + 1) And (h.cap - 1)
    Loop
    HSlot = -1
End Function

Private Function HHas(h As THash, ByVal key As String) As Boolean
    HHas = (HSlot(h, key) >= 0)
End Function

Private Function HGet(h As THash, ByVal key As String) As Long
    Dim i As Long
    i = HSlot(h, key)
    If i < 0 Then Fail "internal: key not found: " & key
    HGet = h.v(i)
End Function

Private Sub HGrow(h As THash)
    Dim ok() As String, ov() As Long, ou() As Byte, oc As Long, i As Long, j As Long
    ok = h.k: ov = h.v: ou = h.u: oc = h.cap
    h.cap = oc * 2: h.n = 0
    ReDim h.k(0 To h.cap - 1): ReDim h.v(0 To h.cap - 1): ReDim h.u(0 To h.cap - 1)
    For i = 0 To oc - 1
        If ou(i) <> 0 Then
            j = HashStr(ok(i)) And (h.cap - 1)
            Do While h.u(j) <> 0
                j = (j + 1) And (h.cap - 1)
            Loop
            h.k(j) = ok(i): h.v(j) = ov(i): h.u(j) = 1: h.n = h.n + 1
        End If
    Next
End Sub

Private Sub HSet(h As THash, ByVal key As String, ByVal v As Long)
    Dim i As Long
    i = HSlot(h, key)
    If i >= 0 Then h.v(i) = v: Exit Sub
    If (h.n + 1) * 2 > h.cap Then HGrow h
    i = HashStr(key) And (h.cap - 1)
    Do While h.u(i) <> 0
        i = (i + 1) And (h.cap - 1)
    Loop
    h.k(i) = key: h.v(i) = v: h.u(i) = 1: h.n = h.n + 1
End Sub

'==============================================================================
' console output (works with the GUI subsystem - own console window - and with a
' console-subsystem build, see MakeConsole.cmd)
'==============================================================================
Private Sub ConInit()
    Dim mode As Long
    ownConsole = (AllocConsole() <> 0)             'fails when a console already exists (console subsystem)
    hOut = GetStdHandle(STD_OUTPUT_HANDLE)
    If ownConsole Then SetConsoleTitle "VBStatic"
    mode = 3                                       'ENABLE_PROCESSED_OUTPUT Or ENABLE_WRAP_AT_EOL_OUTPUT
    outIsConsole = (SetConsoleMode(hOut, mode) <> 0)   'fails when stdout is redirected to a file / pipe
End Sub

Private Sub ConPrint(ByVal s As String, Optional ByVal attr As Long = 7)
    Dim wr As Long
    If hOut = 0 Or hOut = -1 Then Exit Sub
    s = s & vbCrLf
    If outIsConsole Then
        SetConsoleTextAttribute hOut, attr
        WriteConsole hOut, s, Len(s), wr, ByVal 0&
        SetConsoleTextAttribute hOut, 7
    Else
        WriteFile hOut, ByVal s, Len(s), wr, ByVal 0&
    End If
End Sub

Private Sub ConDone()
    If ownConsole Then FreeConsole                  'no pause: the window closes when the tool ends
End Sub

Private Sub Fail(ByVal msg As String)
    Err.Raise vbObjectError + 1, "VBStatic", msg
End Sub

Private Sub AddLog(ByVal s As String)
    If Left$(s, 7) = "WARNING" Then ConPrint s, 14 Else ConPrint s
End Sub

'==============================================================================
' DLL parsing
'==============================================================================
Private Function R2O(ByVal rva As Long) As Long
    Dim i As Long, sz As Long
    For i = 0 To nS - 1
        sz = sVS(i)
        If sRS(i) > sz Then sz = sRS(i)
        If rva >= sVA(i) And rva < sVA(i) + sz Then
            R2O = sRP(i) + rva - sVA(i)
            Exit Function
        End If
    Next
    Fail "RVA &H" & Hex$(rva) & " is not inside any DLL section"
End Function

Private Sub LoadDll(ByVal path As String)
    Dim pe As Long, opt As Long, osz As Long, i As Long, h As Long
    dll = ReadAll(path)
    If W16(dll, 0) <> &H5A4D Then Fail "Not an MZ file: " & path
    pe = L32(dll, &H3C)
    If L32(dll, pe) <> &H4550 Then Fail "No PE signature"
    If W16(dll, pe + 4) <> &H14C Then Fail "Not an i386 PE"
    nS = W16(dll, pe + 6): osz = W16(dll, pe + 20): opt = pe + 24
    If W16(dll, opt) <> &H10B Then Fail "Not a PE32 image"
    If nS > 64 Then Fail "Too many sections"
    epRva = L32(dll, opt + 16): imgBase = L32(dll, opt + 28)
    For i = 0 To 15
        ddRVA(i) = L32(dll, opt + 96 + i * 8): ddSz(i) = L32(dll, opt + 100 + i * 8)
    Next
    For i = 0 To nS - 1
        h = opt + osz + i * 40
        sName(i) = CStrAt(dll, h, 8)
        sVS(i) = L32(dll, h + 8): sVA(i) = L32(dll, h + 12)
        sRS(i) = L32(dll, h + 16): sRP(i) = L32(dll, h + 20)
    Next

    '--- exports
    Dim eo As Long, nf As Long, nn As Long, af As Long, an As Long, ao As Long, nm As String, idx As Long
    If ddSz(0) = 0 Then Fail "DLL has no export table"
    eo = R2O(ddRVA(0))
    expBase = L32(dll, eo + 16): nf = L32(dll, eo + 20): nn = L32(dll, eo + 24)
    af = L32(dll, eo + 28): an = L32(dll, eo + 32): ao = L32(dll, eo + 36)
    expCnt = nf
    ReDim expRva(0 To nf - 1): ReDim expName(0 To nf - 1)
    For i = 0 To nf - 1
        expRva(i) = L32(dll, R2O(af) + 4 * i)
        If expRva(i) >= ddRVA(0) And expRva(i) < ddRVA(0) + ddSz(0) Then Fail "Forwarded exports are not supported"
    Next
    HInit hNames
    For i = 0 To nn - 1
        nm = CStrAt(dll, R2O(L32(dll, R2O(an) + 4 * i)))
        idx = W16(dll, R2O(ao) + 2 * i)
        expName(idx) = nm
        HSet hNames, nm, idx + expBase
    Next

    '--- imports
    Dim io As Long, oft As Long, nmR As Long, ft As Long, t As Long, v As Long, ho As Long
    io = R2O(ddRVA(1))
    Do
        oft = L32(dll, io): nmR = L32(dll, io + 12): ft = L32(dll, io + 16)
        If nmR = 0 Then Exit Do
        nImp = nImp + 1
        If nImp > 64 Then Fail "Too many import descriptors"
        impDll(nImp) = CStrAt(dll, R2O(nmR)): impFT(nImp) = ft: impFirst(nImp) = nEnt + 1
        t = oft
        If t = 0 Then t = ft
        Do
            v = L32(dll, R2O(t)): t = t + 4
            If v = 0 Then Exit Do
            nEnt = nEnt + 1
            If nEnt > 16384 Then Fail "Too many imports"
            If v < 0 Then
                entOrd(nEnt) = v And &HFFFF&: entHint(nEnt) = 0: entName(nEnt) = ""
            Else
                ho = R2O(v)
                entOrd(nEnt) = -1: entHint(nEnt) = W16(dll, ho): entName(nEnt) = CStrAt(dll, ho + 2)
            End If
        Loop
        impCnt(nImp) = nEnt - impFirst(nImp) + 1
        io = io + 20
    Loop

    '--- base relocations
    Dim ro As Long, rend As Long, pg As Long, bs As Long, w As Long, k As Long, cap As Long
    cap = 65536: ReDim relRva(0 To cap - 1)
    ro = R2O(ddRVA(5)): rend = ro + ddSz(5)
    Do While ro < rend
        pg = L32(dll, ro): bs = L32(dll, ro + 4)
        If bs < 8 Then Exit Do
        For k = 0 To (bs - 8) \ 2 - 1
            w = W16(dll, ro + 8 + 2 * k)
            Select Case w \ &H1000&
                Case 0
                Case 3
                    If nRel >= cap Then cap = cap * 2: ReDim Preserve relRva(0 To cap - 1)
                    relRva(nRel) = pg + (w And &HFFF&): nRel = nRel + 1
                Case Else
                    Fail "Unsupported relocation type " & (w \ &H1000&)
            End Select
        Next
        ro = ro + bs
    Loop
    AddLog "DLL: " & nS & " sections, " & nbExports() & " exports, " & nEnt & " imports in " & nImp & " DLLs, " & nRel & " relocations"
End Sub

Private Function nbExports() As Long
    Dim i As Long, c As Long
    For i = 0 To expCnt - 1
        If expRva(i) <> 0 Then c = c + 1
    Next
    nbExports = c
End Function

Private Function FindDllSec(ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To nS - 1
        If sName(i) = nm Then FindDllSec = i: Exit Function
    Next
    Fail "DLL section " & nm & " not found"
End Function

'==============================================================================
' out1.OBJ parsing
'==============================================================================
Private Sub ParseObj(ByVal path As String)
    Dim i As Long, eo As Long, na As Long
    ob = ReadAll(path)
    If W16(ob, 0) <> &H14C Then Fail "Not an i386 COFF object: " & path
    oNs = W16(ob, 2): oSymp = L32(ob, 8): oNsym = L32(ob, 12)
    oStOff = oSymp + oNsym * 18: oStLen = L32(ob, oStOff)
    ReDim oNm(0 To oNsym - 1): ReDim oVal(0 To oNsym - 1): ReDim oSec(0 To oNsym - 1): ReDim oCls(0 To oNsym - 1)
    HInit oHave: HInit oDefined
    i = 0
    Do While i < oNsym
        eo = oSymp + i * 18
        If L32(ob, eo) = 0 Then
            oNm(i) = CStrAt(ob, oStOff + L32(ob, eo + 4))
        Else
            oNm(i) = CStrAt(ob, eo, 8)
        End If
        oVal(i) = L32(ob, eo + 8)
        oSec(i) = W16(ob, eo + 12): If oSec(i) > 32767 Then oSec(i) = oSec(i) - 65536
        oCls(i) = ob(eo + 16): na = ob(eo + 17)
        HSet oHave, oNm(i), i
        If oCls(i) = 2 And (oSec(i) > 0 Or oSec(i) = -1 Or oVal(i) <> 0) Then HSet oDefined, oNm(i), 1
        i = i + 1 + na
    Loop
    AddLog "OBJ: " & oNs & " sections, " & oNsym & " symbol entries, " & oDefined.n & " external definitions"
End Sub

Private Function FindObjSec(ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To oNs - 1
        If CStrAt(ob, 20 + i * 40, 8) = nm Then FindObjSec = i: Exit Function
    Next
    FindObjSec = -1
End Function

'==============================================================================
' resources
'==============================================================================
Private Function ResKey(ByVal c As Long) As String
    If rn(c).isNamed Then ResKey = rn(c).nameStr Else ResKey = CStr(rn(c).id)
End Function

Private Function NewResNode() As Long
    rnN = rnN + 1
    If rnN > 4096 Then Fail "Too many resource nodes"
    NewResNode = rnN
End Function

Private Function ParseResDir(src() As Byte, ByVal base As Long, ByVal off As Long, ByVal fromDll As Boolean) As Long
    Dim n As Long, nn As Long, ni As Long, i As Long, j As Long
    Dim nm As Long, o As Long, c As Long, prev As Long, so As Long, l As Long, s As String
    n = NewResNode()
    rn(n).isDir = True: rn(n).srcDll = fromDll
    nn = W16(src, base + off + 12): ni = W16(src, base + off + 14)
    For i = 0 To nn + ni - 1
        nm = L32(src, base + off + 16 + 8 * i)
        o = L32(src, base + off + 20 + 8 * i)
        If o < 0 Then
            c = ParseResDir(src, base, o And &H7FFFFFFF, fromDll)
        Else
            c = NewResNode()
            rn(c).srcDll = fromDll
            rn(c).dSize = L32(src, base + o + 4)
            rn(c).cp = L32(src, base + o + 8)
            If fromDll Then
                rn(c).dOff = R2O(L32(src, base + o))
            Else
                rn(c).dOff = HGet(hLeaf, CStr(o))
            End If
        End If
        If nm < 0 Then
            so = base + (nm And &H7FFFFFFF): l = W16(src, so): s = ""
            For j = 0 To l - 1
                s = s & ChrW$(W16(src, so + 2 + 2 * j))
            Next
            rn(c).isNamed = True: rn(c).nameStr = s
        Else
            rn(c).id = nm
        End If
        If prev = 0 Then rn(n).child = c Else rn(prev).sib = c
        prev = c
        rn(n).nChild = rn(n).nChild + 1
    Next
    ParseResDir = n
End Function

Private Function FindResChild(ByVal a As Long, ByVal cb As Long) As Long
    Dim c As Long
    c = rn(a).child
    Do While c <> 0
        If rn(c).isNamed = rn(cb).isNamed Then
            If rn(c).isNamed Then
                If StrComp(rn(c).nameStr, rn(cb).nameStr, vbTextCompare) = 0 Then FindResChild = c: Exit Function
            ElseIf rn(c).id = rn(cb).id Then
                FindResChild = c: Exit Function
            End If
        End If
        c = rn(c).sib
    Loop
End Function

Private Sub MergeRes(ByVal a As Long, ByVal b As Long, ByVal path As String)
    Dim cb As Long, nx As Long, ca As Long, tail As Long, p As String
    cb = rn(b).child
    Do While cb <> 0
        nx = rn(cb).sib: rn(cb).sib = 0
        ca = FindResChild(a, cb)
        p = path & "/" & ResKey(cb)
        If ca = 0 Then
            If rn(a).child = 0 Then
                rn(a).child = cb
            Else
                tail = rn(a).child
                Do While rn(tail).sib <> 0
                    tail = rn(tail).sib
                Loop
                rn(tail).sib = cb
            End If
            rn(a).nChild = rn(a).nChild + 1
        ElseIf rn(ca).isDir And rn(cb).isDir Then
            MergeRes ca, cb, p
        Else
            resConflicts = resConflicts & p & " "
        End If
        cb = nx
    Loop
End Sub

Private Function ResLess(ByVal x As Long, ByVal y As Long) As Boolean
    If rn(x).isNamed <> rn(y).isNamed Then
        ResLess = rn(x).isNamed
    ElseIf rn(x).isNamed Then
        ResLess = (StrComp(rn(x).nameStr, rn(y).nameStr, vbTextCompare) < 0)
    Else
        ResLess = (rn(x).id < rn(y).id)
    End If
End Function

Private Function KidsSorted(ByVal n As Long, kids() As Long) As Long
    Dim c As Long, k As Long, i As Long, j As Long, t As Long
    ReDim kids(0 To rn(n).nChild)
    c = rn(n).child
    Do While c <> 0
        kids(k) = c: k = k + 1
        c = rn(c).sib
    Loop
    For i = 1 To k - 1
        t = kids(i): j = i - 1
        Do While j >= 0
            If ResLess(t, kids(j)) Then
                kids(j + 1) = kids(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        kids(j + 1) = t
    Next
    KidsSorted = k
End Function

Private Sub AssignDirs(ByVal n As Long)
    Dim kids() As Long, k As Long, i As Long
    rn(n).off = rCur
    rCur = rCur + 16 + 8 * rn(n).nChild
    k = KidsSorted(n, kids)
    For i = 0 To k - 1
        If rn(kids(i)).isNamed Then
            If Not HHas(hStr, rn(kids(i)).nameStr) Then HSet hStr, rn(kids(i)).nameStr, -1
        End If
        If rn(kids(i)).isDir Then
            AssignDirs kids(i)
        Else
            rLeafN = rLeafN + 1: rLeaf(rLeafN) = kids(i)
        End If
    Next
End Sub

Private Sub WriteResDir(ByVal n As Long)
    Dim kids() As Long, k As Long, i As Long, o As Long, f As Long, t As Long, nNamed As Long
    k = KidsSorted(n, kids)
    o = rn(n).off
    For i = 0 To k - 1
        If rn(kids(i)).isNamed Then nNamed = nNamed + 1
    Next
    BPatchW rsr01, o + 12, nNamed
    BPatchW rsr01, o + 14, k - nNamed
    For i = 0 To k - 1
        If rn(kids(i)).isNamed Then f = HGet(hStr, rn(kids(i)).nameStr) Or &H80000000 Else f = rn(kids(i)).id
        If rn(kids(i)).isDir Then t = rn(kids(i)).off Or &H80000000 Else t = rn(kids(i)).off
        BPatchL rsr01, o + 16 + 8 * i, f
        BPatchL rsr01, o + 20 + 8 * i, t
        If rn(kids(i)).isDir Then WriteResDir kids(i)
    Next
End Sub

Private Sub LayoutResources(ByVal root As Long)
    Dim blank As TBuf, i As Long, c As Long, s As String, j As Long, o As Long
    rsr01 = blank: rsr02 = blank: rsrRel = blank
    HInit hStr
    rCur = 0: rLeafN = 0
    AssignDirs root
    For i = 1 To rLeafN
        rn(rLeaf(i)).off = rCur: rCur = rCur + 16
    Next
    For i = 0 To hStr.cap - 1
        If hStr.u(i) <> 0 Then
            hStr.v(i) = rCur
            rCur = rCur + ((2 + 2 * Len(hStr.k(i)) + 3) \ 4) * 4
        End If
    Next
    BZeros rsr01, rCur
    WriteResDir root
    For i = 0 To hStr.cap - 1                         'name strings (UTF-16, length prefixed)
        If hStr.u(i) <> 0 Then
            s = hStr.k(i): o = hStr.v(i)
            BPatchW rsr01, o, Len(s)
            For j = 1 To Len(s)
                BPatchW rsr01, o + 2 * j, AscW(Mid$(s, j, 1)) And &HFFFF&
            Next
        End If
    Next
    For i = 1 To rLeafN                               'data entries + data blobs
        c = rLeaf(i)
        BPad rsr02, 4
        BPatchL rsr01, rn(c).off + 0, rsr02.n
        BPatchL rsr01, rn(c).off + 4, rn(c).dSize
        BPatchL rsr01, rn(c).off + 8, rn(c).cp
        If rn(c).srcDll Then
            BPutBytes rsr02, dll, rn(c).dOff, rn(c).dSize
        Else
            BPutBytes rsr02, ob, rn(c).dOff, rn(c).dSize
        End If
        BPutL rsrRel, rn(c).off
    Next
    rsr01.n = rCur
End Sub

Private Function CountLeaves(ByVal n As Long) As Long
    Dim c As Long, t As Long
    If Not rn(n).isDir Then CountLeaves = 1: Exit Function
    c = rn(n).child
    Do While c <> 0
        t = t + CountLeaves(c)
        c = rn(c).sib
    Loop
    CountLeaves = t
End Function

Private Function FindById(ByVal a As Long, ByVal id As Long) As Long
    Dim c As Long
    c = rn(a).child
    Do While c <> 0
        If (Not rn(c).isNamed) And rn(c).id = id Then FindById = c: Exit Function
        c = rn(c).sib
    Loop
End Function

'/NOICON: msvbvm60 loads its own icons (RT_GROUP_ICON 1201-1204 + RT_ICON 1-11) with LoadImageA from
'its module, so they MUST stay in the output. Explorer, however, shows the first RT_GROUP_ICON of
'the exe (named entries first, then the lowest ID). When the OBJ has no icon in front of the
'runtime's, an empty group icon (ID 1, 0 images) is added: Explorer cannot extract anything from
'it and shows the generic application icon; LoadImageA for 1201-1204 is not affected.
Private Sub HideRuntimeIcons(ByVal root As Long)
    Dim t As Long, kids() As Long, k As Long, i As Long, nm As Long, lf As Long, tail As Long, n0 As Long
    t = FindById(root, 14)
    If t = 0 Then AddLog "resources: /NOICON - no RT_GROUP_ICON present, nothing to hide": Exit Sub
    k = KidsSorted(t, kids)
    If k = 0 Then Exit Sub
    If Not rn(kids(0)).srcDll Then AddLog "resources: /NOICON - the first icon group is the OBJ's own, nothing to hide": Exit Sub
    For i = 0 To k - 1
        If Not rn(kids(i)).srcDll Then
            AddLog "resources: /NOICON WARNING - the OBJ's icon group has a higher ID/name than the runtime's (" & rn(kids(0)).id & "); Explorer will show the runtime icon. Give the app icon an ID below 1201."
            Exit Sub
        End If
    Next
    If rn(kids(0)).isNamed Or rn(kids(0)).id <= 1 Then AddLog "resources: /NOICON - cannot place an empty icon group in front, skipped": Exit Sub
    n0 = UBound(ob) + 1                               '6 bytes ICONDIR: reserved 0, type 1, count 0
    ReDim Preserve ob(0 To n0 + 7)
    ob(n0 + 2) = 1
    nm = NewResNode(): rn(nm).isDir = True: rn(nm).id = 1
    lf = NewResNode(): rn(lf).id = 0: rn(lf).dSize = 6: rn(lf).dOff = n0
    rn(nm).child = lf: rn(nm).nChild = 1
    tail = rn(t).child
    Do While rn(tail).sib <> 0
        tail = rn(tail).sib
    Loop
    rn(tail).sib = nm: rn(t).nChild = rn(t).nChild + 1
    AddLog "resources: /NOICON - empty RT_GROUP_ICON 1 added in front of the runtime's icons (they stay for LoadImage)"
End Sub

Private Sub PlanResources()
    Dim i1 As Long, i2 As Long, h1 As Long, h2 As Long, k As Long, root As Long, rootD As Long
    Dim rs1 As Long, rp1 As Long, rr1 As Long, nr1 As Long, rp2 As Long, va As Long, si As Long
    resMode = 0: rnN = 0: resConflicts = ""
    If ddSz(2) = 0 Then AddLog "resources: the DLL has none": Exit Sub
    rootD = ParseResDir(dll, R2O(ddRVA(2)), 0, True)
    i1 = FindObjSec(".rsrc$01"): i2 = FindObjSec(".rsrc$02")
    If i1 >= 0 And i2 >= 0 Then
        h1 = 20 + i1 * 40: h2 = 20 + i2 * 40
        rs1 = L32(ob, h1 + 16): rp1 = L32(ob, h1 + 20): rr1 = L32(ob, h1 + 24): nr1 = W16(ob, h1 + 32)
        rp2 = L32(ob, h2 + 20)
        HInit hLeaf
        For k = 0 To nr1 - 1
            va = L32(ob, rr1 + 10 * k): si = L32(ob, rr1 + 10 * k + 4)
            HSet hLeaf, CStr(va), rp2 + oVal(si) + L32(ob, rp1 + va)
        Next
        root = ParseResDir(ob, rp1, 0, False)
        MergeRes root, rootD, ""
        resMode = 1
    Else
        root = rootD: resMode = 2
    End If
    If gNoIcon Then HideRuntimeIcons root
    LayoutResources root
    If resMode = 1 Then
        AddLog "resources: merged into OBJ (" & rLeafN & " items, " & rsr02.n \ 1024 & " KB)"
    Else
        AddLog "resources: OBJ has none - runtime resources are added as new sections (" & rLeafN & " items)"
    End If
    If Len(resConflicts) > 0 Then AddLog "resource conflicts (OBJ's version kept): " & resConflicts
End Sub

'==============================================================================
' COFF writer
'==============================================================================
Private Function AddSym(ByVal nm As String, ByVal value As Long, ByVal sec As Long, ByVal typ As Long, ByVal cls As Long, Optional ByVal aux As Boolean = False) As Long
    Dim i As Long
    AddSym = symN
    HSet oHave, nm, symN                          'so later lookups see our own symbols too
    If Len(nm) <= 8 Then
        For i = 1 To 8
            If i <= Len(nm) Then BPutB symTab, Asc(Mid$(nm, i, 1)) Else BPutB symTab, 0
        Next
    Else
        BPutL symTab, 0
        BPutL symTab, strTab.n
        BPutStr strTab, nm
        BPutB strTab, 0
    End If
    BPutL symTab, value
    BPutW symTab, sec
    BPutW symTab, typ
    BPutB symTab, cls
    If aux Then BPutB symTab, 1 Else BPutB symTab, 0
    symN = symN + 1
    If aux Then BZeros symTab, 18: symN = symN + 1
End Function

'reference to a global symbol: the OBJ's own entry if there is one, else a new undefined external
Private Function ExternSym(ByVal nm As String) As Long
    If HHas(oHave, nm) Then ExternSym = HGet(oHave, nm) Else ExternSym = AddSym(nm, 0, 0, 0, 2)
End Function

'define a global symbol: an undefined entry of the OBJ is filled in place, otherwise a new one is added
Private Function DefineSym(ByVal nm As String, ByVal value As Long, ByVal sec As Long, ByVal typ As Long) As Long
    Dim i As Long, o As Long
    If HHas(oHave, nm) Then
        i = HGet(oHave, nm): o = i * 18
        If Not (symTab.d(o + 16) = 2 And BGetL(symTab, o + 8) = 0 And TW16(symTab, o + 12) = 0) Then Fail "Duplicate definition of symbol " & nm
        BPatchL symTab, o + 8, value
        BPatchW symTab, o + 12, sec
        BPatchW symTab, o + 14, typ
        DefineSym = i
    Else
        DefineSym = AddSym(nm, value, sec, typ, 2)
    End If
End Function

Private Function NewSec(ByVal nm As String, ByVal flags As Long, Optional ByVal initSize As Long = 0) As Long
    csN = csN + 1
    If csN > 90 Then Fail "internal: too many sections"
    cs(csN).nm = nm: cs(csN).flags = flags
    If initSize > 0 Then BZeros cs(csN).data, initSize
    cs(csN).symIdx = AddSym(nm, 0, csN, 0, 3, True)
    NewSec = csN
End Function

Private Sub AddRel(ByVal s As Long, ByVal va As Long, ByVal sym As Long, ByVal typ As Long)
    If cs(s).nrel >= cs(s).relcap Then
        cs(s).relcap = cs(s).relcap * 2 + 1024
        If cs(s).nrel = 0 Then
            ReDim cs(s).rel(0 To cs(s).relcap - 1)
        Else
            ReDim Preserve cs(s).rel(0 To cs(s).relcap - 1)
        End If
    End If
    cs(s).rel(cs(s).nrel).va = va
    cs(s).rel(cs(s).nrel).sym = sym
    cs(s).rel(cs(s).nrel).typ = typ
    cs(s).nrel = cs(s).nrel + 1
End Sub

Private Sub Target(ByVal rva As Long, ByRef sec As Long, ByRef off As Long)
    If rva >= iatRva And rva < iatRva + iatSz Then sec = sD5: off = rva - iatRva: Exit Sub
    If rva >= codeLo And rva <= codeHi Then sec = sA: off = rva - codeLo: Exit Sub
    If rva >= dVA And rva <= dVA + dVS Then sec = sB: off = rva - dVA: Exit Sub
    Fail "Relocation target &H" & Hex$(rva) & " is outside the embedded sections (resources/headers?)"
End Sub

Private Sub BuildObjFile(outB As TBuf)
    Dim hdr As Long, body As TBuf, heads As TBuf
    Dim i As Long, k As Long, dp As Long, rp As Long, o As Long
    ' fix section-symbol aux records (length, reloc count)
    For i = 1 To csN
        If cs(i).symIdx >= 0 Then
            o = cs(i).symIdx * 18 + 18
            If cs(i).bss > 0 Then BPatchL symTab, o, cs(i).bss Else BPatchL symTab, o, cs(i).data.n
            symTab.d(o + 4) = CByte(cs(i).nrel And &HFF&)
            symTab.d(o + 5) = CByte((cs(i).nrel And &HFF00&) \ &H100&)
        End If
    Next
    hdr = 20 + 40 * csN
    For i = 1 To csN
        dp = 0: rp = 0
        If cs(i).data.n > 0 Then
            dp = hdr + body.n
            BPutBytes body, cs(i).data.d, 0, cs(i).data.n
            BPad body, 4
        End If
        If cs(i).nrel > 65534 Then Fail "Section " & cs(i).nm & " has too many relocations"
        If cs(i).nrel > 0 Then
            rp = hdr + body.n
            For k = 0 To cs(i).nrel - 1
                BPutL body, cs(i).rel(k).va
                BPutL body, cs(i).rel(k).sym
                BPutW body, cs(i).rel(k).typ
            Next
            BPad body, 4
        End If
        For k = 1 To 8
            If k <= Len(cs(i).nm) Then BPutB heads, Asc(Mid$(cs(i).nm, k, 1)) Else BPutB heads, 0
        Next
        BPutL heads, 0: BPutL heads, 0
        If cs(i).bss > 0 Then BPutL heads, cs(i).bss Else BPutL heads, cs(i).data.n
        BPutL heads, dp
        BPutL heads, rp
        BPutL heads, 0
        BPutW heads, cs(i).nrel
        BPutW heads, 0
        BPutL heads, cs(i).flags
    Next
    BPutW outB, &H14C: BPutW outB, csN: BPutL outB, 0
    BPutL outB, hdr + body.n: BPutL outB, symN: BPutW outB, 0: BPutW outB, 0
    BPutBytes outB, heads.d, 0, heads.n
    BPutBytes outB, body.d, 0, body.n
    BPutBytes outB, symTab.d, 0, symTab.n
    BPatchL strTab, 0, strTab.n
    BPutBytes outB, strTab.d, 0, strTab.n
End Sub

'==============================================================================
' load the main OBJ into the writer (sections, relocations, symbols, strings)
'==============================================================================
Private Sub ImportMain()
    Dim blank As TBuf, i As Long, h As Long, k As Long, rs As Long, rp As Long, rr As Long, nr As Long, na As Long
    symTab = blank: strTab = blank: symN = 0: csN = 0
    ReDim cs(1 To 96)
    If oNs > 60 Then Fail "The OBJ has too many sections"
    For i = 1 To oNs
        h = 20 + (i - 1) * 40
        csN = csN + 1
        cs(i).nm = CStrAt(ob, h, 8)
        rs = L32(ob, h + 16): rp = L32(ob, h + 20): rr = L32(ob, h + 24): nr = W16(ob, h + 32)
        If L32(ob, h + 28) <> 0 Or W16(ob, h + 34) <> 0 Then Fail "Line numbers in section " & cs(i).nm & " are not supported"
        cs(i).flags = L32(ob, h + 36)
        cs(i).symIdx = -1
        If rp <> 0 Then BPutBytes cs(i).data, ob, rp, rs Else cs(i).bss = rs
        For k = 0 To nr - 1
            AddRel i, L32(ob, rr + 10 * k), L32(ob, rr + 10 * k + 4), W16(ob, rr + 10 * k + 8)
        Next
    Next
    BPutBytes symTab, ob, oSymp, oNsym * 18          'symbol table and string table are taken over verbatim,
    symN = oNsym                                     'so every symbol index in the relocations stays valid
    BPutBytes strTab, ob, oStOff, oStLen
    i = 0
    Do While i < oNsym                               'find the section symbol of every section (for aux fix-ups)
        na = ob(oSymp + i * 18 + 17)
        If oCls(i) = 3 And na = 1 And oSec(i) > 0 And oVal(i) = 0 Then
            If cs(oSec(i)).symIdx < 0 Then cs(oSec(i)).symIdx = i
        End If
        i = i + 1 + na
    Loop
End Sub

'==============================================================================
' add the runtime to the writer
'==============================================================================
Private Sub BuildVbrt(ByVal chainSym As String, ByVal entryName As String, ByVal dllMode As Boolean)
    Dim i As Long, k As Long, e As Long, o As Long, l As Long
    Dim iT As Long, iE As Long, iD As Long

    iT = FindDllSec(".text"): iE = FindDllSec("ENGINE"): iD = FindDllSec(".data")
    codeLo = sVA(iT): codeHi = sVA(iE) + sVS(iE)
    dVA = sVA(iD): dVS = sVS(iD)
    iatRva = ddRVA(12): iatSz = ddSz(12)
    If iatSz = 0 Then Fail "DLL has no IAT directory"

    sA = NewSec(".vbrt", &H60D00020)                 'code, 4K aligned
    sB = NewSec(".vbdat", &HC0D00040)                'data rw, 4K aligned
    sT = NewSec(".vbthk", &H60500020)                'import-lib style thunks
    sM = NewSec(".vbimp", &HC0500040)                '__imp_ cells
    sD2 = NewSec(".idata$2", &HC0300040, 20 * nImp)
    sD3 = NewSec(".idata$3", &HC0300040, 20)             'import directory terminator
    sD4 = NewSec(".idata$4", &HC0300040, iatSz)
    sD5 = NewSec(".idata$5", &HC0300040, iatSz)
    sD6 = NewSec(".idata$6", &HC0300040)
    sE = NewSec(".vbent", &H60500020)

    '--- section images
    BZeros cs(sA).data, codeHi - codeLo
    For k = 0 To 1
        If k = 0 Then e = iT Else e = iE
        l = sVS(e): If sRS(e) < l Then l = sRS(e)
        CopyMemory cs(sA).data.d(sVA(e) - codeLo), dll(sRP(e)), l
    Next
    BZeros cs(sB).data, dVS
    l = sVS(iD): If sRS(iD) < l Then l = sRS(iD)
    CopyMemory cs(sB).data.d(0), dll(sRP(iD)), l

    '--- DLL relocations -> COFF DIR32
    Dim r As Long, sc As Long, ts As Long, toff As Long, v As Long, skipped As Long
    For k = 0 To nRel - 1
        r = relRva(k)
        sc = 0
        If r >= codeLo And r < codeHi Then
            sc = sA: o = r - codeLo
        ElseIf r >= dVA And r < dVA + dVS Then
            sc = sB: o = r - dVA
        End If
        If sc = 0 Then
            skipped = skipped + 1
        Else
            v = BGetL(cs(sc).data, o) - imgBase
            Target v, ts, toff
            BPatchL cs(sc).data, o, toff
            AddRel sc, o, cs(ts).symIdx, R_DIR32
        End If
    Next
    If skipped > 0 Then AddLog "WARNING: " & skipped & " relocations lie outside .text/ENGINE/.data (ignored)"

    '--- rebuild import directory
    Dim nameOff() As Long, ho As Long, base As Long, d As Long
    Dim hSlot As THash: HInit hSlot
    ReDim nameOff(1 To nImp)
    For d = 1 To nImp
        nameOff(d) = cs(sD6).data.n
        BPutStr cs(sD6).data, impDll(d): BPutB cs(sD6).data, 0
        BPad cs(sD6).data, 2
    Next
    For d = 1 To nImp
        base = impFT(d) - iatRva
        For k = 0 To impCnt(d) - 1
            e = impFirst(d) + k: o = base + 4 * k
            If entOrd(e) >= 0 Then
                BPatchL cs(sD4).data, o, &H80000000 Or entOrd(e)
                BPatchL cs(sD5).data, o, &H80000000 Or entOrd(e)
            Else
                ho = cs(sD6).data.n
                BPutW cs(sD6).data, entHint(e)
                BPutStr cs(sD6).data, entName(e): BPutB cs(sD6).data, 0
                BPad cs(sD6).data, 2
                BPatchL cs(sD4).data, o, ho
                BPatchL cs(sD5).data, o, ho
                AddRel sD4, o, cs(sD6).symIdx, R_DIR32NB
                AddRel sD5, o, cs(sD6).symIdx, R_DIR32NB
                HSet hSlot, UCase$(impDll(d)) & "|" & entName(e), o
            End If
        Next
        o = 20 * (d - 1)
        BPatchL cs(sD2).data, o + 0, base
        BPatchL cs(sD2).data, o + 12, nameOff(d)
        BPatchL cs(sD2).data, o + 16, base
        AddRel sD2, o + 0, cs(sD4).symIdx, R_DIR32NB
        AddRel sD2, o + 12, cs(sD6).symIdx, R_DIR32NB
        AddRel sD2, o + 16, cs(sD5).symIdx, R_DIR32NB
    Next

    Dim rva As Long

    '--- import-library style  _Name  /  __imp__Name  for names out1 doesn't provide
    Dim nm As String, sy As String, impSym As String, ci As Long, nThunks As Long
    Dim needCell As Boolean, needThunk As Boolean
    For i = 0 To expCnt - 1
        nm = expName(i)
        rva = expRva(i)
        If Len(nm) > 0 And rva >= codeLo And rva < codeHi Then
            sy = "_" & nm: impSym = "__imp_" & sy
            needCell = Not HHas(oDefined, impSym)
            needThunk = Not HHas(oDefined, sy)
            If needCell Or needThunk Then
                If needCell Then
                    o = cs(sM).data.n
                    BPutL cs(sM).data, rva - codeLo
                    AddRel sM, o, cs(sA).symIdx, R_DIR32
                    ci = DefineSym(impSym, o, sM, 0)
                Else
                    ci = ExternSym(impSym)
                End If
                If needThunk Then
                    o = cs(sT).data.n
                    BPutB cs(sT).data, &HFF: BPutB cs(sT).data, &H25: BPutL cs(sT).data, 0
                    BPutB cs(sT).data, &H90: BPutB cs(sT).data, &H90
                    AddRel sT, o + 2, ci, R_DIR32
                    DefineSym sy, o, sT, &H20&
                    nThunks = nThunks + 1
                End If
            End If
        End If
    Next
    AddLog "vbrt: " & nThunks & " extra import thunks generated"

    '--- symbols normally supplied by msvbvm60.lib / VBAEXE6.LIB(natsupp.obj)
    '    natsupp.obj defines exactly these three globals. If the patched OBJ defined only some of
    '    them, every other OBJ that still references the missing one (typically __fltused) makes
    '    the linker pull natsupp.obj out of VBAEXE6.LIB, and its __adjust_fdiv then collides with
    '    ours (LNK2005 / LNK1169). So all three are defined here -> natsupp.obj is never extracted.
    '    (__fltused / __except_list are absolute symbols, so they never clash with the lib's copy.)
    If Not HHas(oDefined, "__except_list") Then DefineSym "__except_list", 0, -1, 0
    If Not HHas(oDefined, "__fltused") Then DefineSym "__fltused", &H9876&, -1, 0
    If Not HHas(oDefined, "__adjust_fdiv") Then
        o = cs(sM).data.n: BPutL cs(sM).data, 0
        DefineSym "__adjust_fdiv", o, sM, 0
    End If
    AddLog "vbrt: __adjust_fdiv / __fltused / __except_list defined (VBAEXE6.LIB(natsupp.obj) is not needed)"

    '--- entry stub: DllMain(GetModuleHandleA(0), DLL_PROCESS_ATTACH, 0) then jmp chainSym
    Dim gm As Long
    If Not HHas(hSlot, "KERNEL32.DLL|GetModuleHandleA") Then Fail "DLL does not import KERNEL32!GetModuleHandleA"
    gm = HGet(hSlot, "KERNEL32.DLL|GetModuleHandleA")
    If Not dllMode Then
        With cs(sE)
            BPutB .data, &H6A: BPutB .data, 0                       'push 0
            BPutB .data, &HFF: BPutB .data, &H15: BPutL .data, gm   'call [GetModuleHandleA]
            BPutB .data, &H6A: BPutB .data, 0                       'push 0   (reserved)
            BPutB .data, &H6A: BPutB .data, 1                       'push 1   (DLL_PROCESS_ATTACH)
            BPutB .data, &H50                                       'push eax (hInstance)
            BPutB .data, &HB8: BPutL .data, epRva - codeLo          'mov eax, DllMain
            BPutB .data, &HFF: BPutB .data, &HD0                    'call eax
            BPutB .data, &H85: BPutB .data, &HC0                    'test eax,eax
            BPutB .data, &H74: BPutB .data, 5                       'jz  fail
            BPutB .data, &HE9: BPutL .data, 0                       'jmp chainSym
            BPutB .data, &H31: BPutB .data, &HC0                    'fail: xor eax,eax
            BPutB .data, &HC3                                       'ret
        End With
        AddRel sE, 4, cs(sD5).symIdx, R_DIR32
        AddRel sE, 14, cs(sA).symIdx, R_DIR32
        DefineSym entryName, 0, sE, &H20&
        ci = ExternSym(chainSym)
        AddRel sE, 25, ci, R_REL32
    Else
        'stdcall DllEntry(hInst, reason, reserved):
        '   reason = 0 (detach) or 3: project entry first, then runtime DllMain
        '   otherwise:                runtime DllMain first (return 0 on failure), then project entry
        Dim a3 As String
        a3 = "FF 75 10 FF 75 0C FF 75 08"
        With cs(sE)
            BPutHex .data, "55 89 E5 83 7D 0C 00 74 2A 83 7D 0C 03 74 24"       '  0..14
            BPutHex .data, a3: BPutB .data, &HB8: BPutL .data, epRva - codeLo     ' 15..28  mov eax,DllMain
            BPutHex .data, "FF D0 85 C0 74 30"                                    ' 29..34
            BPutHex .data, a3: BPutB .data, &HE8: BPutL .data, 0                  ' 35..48  call chain
            BPutHex .data, "EB 20"                                                ' 49..50
            BPutHex .data, a3: BPutB .data, &HE8: BPutL .data, 0                  ' 51..64  detach: call chain
            BPutB .data, &H50                                                     ' 65      push eax
            BPutHex .data, a3: BPutB .data, &HB8: BPutL .data, epRva - codeLo     ' 66..79  mov eax,DllMain
            BPutHex .data, "FF D0 58 5D C2 0C 00"                                 ' 80..86
        End With
        If cs(sE).data.n <> 87 Then Fail "internal: DLL entry stub has wrong size"
        ci = ExternSym(chainSym)
        AddRel sE, 25, cs(sA).symIdx, R_DIR32
        AddRel sE, 45, ci, R_REL32
        AddRel sE, 61, ci, R_REL32
        AddRel sE, 76, cs(sA).symIdx, R_DIR32
        DefineSym entryName, 0, sE, &H20&
    End If

    '--- resources of the runtime, when the OBJ has no resource sections of its own
    If resMode = 2 Then
        Dim r1 As Long, r2 As Long
        r1 = NewSec(".rsrc$01", &H40300040)
        r2 = NewSec(".rsrc$02", &H40300040)
        BPutBytes cs(r1).data, rsr01.d, 0, rsr01.n
        BPutBytes cs(r2).data, rsr02.d, 0, rsr02.n
        For k = 0 To rsrRel.n \ 4 - 1
            AddRel r1, BGetL(rsrRel, k * 4), cs(r2).symIdx, R_DIR32NB
        Next
    End If
End Sub

'==============================================================================
' patch the main OBJ's own sections (inside the writer)
'==============================================================================
Private Sub PatchMain()
    Dim i As Long, k As Long, off As Long, h As Long
    Dim i2 As Long, i3 As Long, i4 As Long, i5 As Long, i6 As Long
    i2 = FindObjSec(".idata$2"): i3 = FindObjSec(".idata$3"): i4 = FindObjSec(".idata$4"): i5 = FindObjSec(".idata$5"): i6 = FindObjSec(".idata$6")
    If i2 < 0 Or i4 < 0 Or i5 < 0 Or i6 < 0 Then Fail "This OBJ has no .idata$2/4/5/6 sections - not a VB6 main object (the one with the runtime imports)?"

    Dim rs4 As Long, rp4 As Long, rr4 As Long, nr4 As Long, rp6 As Long
    h = 20 + i4 * 40: rs4 = L32(ob, h + 16): rp4 = L32(ob, h + 20): rr4 = L32(ob, h + 24): nr4 = W16(ob, h + 32)
    rp6 = L32(ob, 20 + i6 * 40 + 20)

    '--- which export does each IAT slot stand for?
    Dim slotOrd() As Long, va As Long, si As Long, nm As String, v As Long, symOff As Long
    ReDim slotOrd(0 To rs4 \ 4 - 1)
    For k = 0 To nr4 - 1
        va = L32(ob, rr4 + 10 * k): si = L32(ob, rr4 + 10 * k + 4)
        symOff = oVal(si) + L32(ob, rp4 + va)
        nm = CStrAt(ob, rp6 + symOff + 2)
        If Not HHas(hNames, nm) Then Fail "The OBJ imports '" & nm & "' which MSVBVM60.DLL does not export"
        slotOrd(va \ 4) = HGet(hNames, nm)
    Next
    For off = 0 To rs4 - 4 Step 4
        v = L32(ob, rp4 + off)
        If v < 0 Then slotOrd(off \ 4) = v And &HFFFF&
    Next

    '--- IAT (.idata$5) becomes plain data: every slot is relocated straight into the embedded runtime
    Dim c5 As Long, rva As Long, nSlots As Long
    c5 = i5 + 1
    cs(c5).nrel = 0
    For k = 0 To UBound(slotOrd)
        If slotOrd(k) > 0 Then
            If slotOrd(k) - expBase < 0 Or slotOrd(k) - expBase >= expCnt Then Fail "Ordinal " & slotOrd(k) & " is not in the DLL"
            rva = expRva(slotOrd(k) - expBase)
            If rva < codeLo Or rva >= codeHi Then Fail "Ordinal " & slotOrd(k) & " is not exported (as code) by the DLL"
            BPatchL cs(c5).data, k * 4, rva - codeLo
            AddRel c5, k * 4, cs(sA).symIdx, R_DIR32
            nSlots = nSlots + 1
        End If
    Next
    cs(c5).nm = ".vbiat"
    If cs(c5).symIdx >= 0 Then SetName8 symTab, cs(c5).symIdx * 18, ".vbiat"

    '--- the old msvbvm60 import sections: emptied AND renamed (.vbnulN), so the linker neither
    '    sees a second import directory nor treats the object as an import-library member
    Dim idxs As Variant, digs As Variant, j As Long
    idxs = Array(i2, i3, i4, i6): digs = Array("2", "3", "4", "6")
    For j = 0 To 3
        If idxs(j) >= 0 Then
            k = idxs(j) + 1
            cs(k).nm = ".vbnul" & digs(j): cs(k).data.n = 0: cs(k).nrel = 0: cs(k).bss = 0
            If cs(k).symIdx >= 0 Then SetName8 symTab, cs(k).symIdx * 18, cs(k).nm
        End If
    Next

    '--- runtime resources merged into the OBJ's own resource sections
    If resMode = 1 Then
        Dim c1 As Long, c2 As Long
        c1 = FindObjSec(".rsrc$01") + 1: c2 = FindObjSec(".rsrc$02") + 1
        If cs(c2).symIdx < 0 Then Fail "No section symbol for .rsrc$02 in the OBJ"
        cs(c1).data.n = 0: BPutBytes cs(c1).data, rsr01.d, 0, rsr01.n: cs(c1).nrel = 0
        For k = 0 To rsrRel.n \ 4 - 1
            AddRel c1, BGetL(rsrRel, k * 4), cs(c2).symIdx, R_DIR32NB
        Next
        cs(c2).data.n = 0: BPutBytes cs(c2).data, rsr02.d, 0, rsr02.n
    End If

    '--- /DLL:STANDARD: drop every -export: directive, ActiveX entry functions become static
    If gStrip Then
        Dim stripped As THash, sk As String, dr As Long, txt As String, tok As Variant, keep As String
        Dim nm2 As String, q As Long, newTxt As String, o As Long
        HInit stripped
        dr = FindObjSec(".drectve")
        If dr >= 0 Then
            k = dr + 1
            If cs(k).data.n > 0 Then txt = CStrAt(cs(k).data.d, 0, cs(k).data.n)
            For Each tok In Split(txt, " ")
                If Len(tok) > 0 Then
                    If LCase$(Left$(tok, 8)) = "-export:" Or LCase$(Left$(tok, 8)) = "/export:" Then
                        nm2 = Mid$(tok, 9)
                        q = InStr(nm2, "=")
                        If q > 0 Then nm2 = Mid$(nm2, q + 1)        'internal name
                        q = InStr(nm2, ",")
                        If q > 0 Then nm2 = Left$(nm2, q - 1)
                        HSet stripped, nm2, 1
                    Else
                        keep = keep & " " & tok
                    End If
                End If
            Next
            If Len(keep) > 0 Then newTxt = keep & " "
            cs(k).data.n = 0
            BPutStr cs(k).data, newTxt
            For i = 0 To stripped.cap - 1
                If stripped.u(i) <> 0 Then
                    sk = sk & " " & stripped.k(i)
                    If HHas(oHave, stripped.k(i)) Then
                        o = HGet(oHave, stripped.k(i)) * 18
                        If symTab.d(o + 16) = 2 And TW16(symTab, o + 12) > 0 And TW16(symTab, o + 12) < 32768 Then symTab.d(o + 16) = 3
                    End If
                End If
            Next
            AddLog "STANDARD DLL: removed " & stripped.n & " export directive(s):" & sk
        Else
            AddLog "STANDARD DLL: OBJ has no .drectve section - nothing to remove"
        End If
    End If
    AddLog "patched IAT: " & nSlots & " slots now relocated into the embedded runtime"
End Sub

'==============================================================================
' command line
'==============================================================================
Private Function SplitArgs(ByVal s As String, args() As String) As Long
    Dim i As Long, c As String, cur As String, inQ As Boolean, n As Long
    ReDim args(0 To 15)
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then c = Mid$(s, i, 1) Else c = " "
        If c = """" Then
            inQ = Not inQ
        ElseIf c = " " And Not inQ Then
            If Len(cur) > 0 Then
                If n < 16 Then args(n) = cur: n = n + 1
                cur = ""
            End If
        Else
            cur = cur & c
        End If
    Next
    SplitArgs = n
End Function

Private Sub RunMain()
    Dim args() As String, n As Long, i As Long, np As Long, pos(0 To 15) As String
    Dim dllPath As String, objPath As String, outPath As String, chain As String
    Dim entryName As String, dllMode As Boolean
    Dim res As TBuf
    On Error GoTo EH
    n = SplitArgs(Command$, args)
    For i = 0 To n - 1
        If Left$(args(i), 1) = "/" Then
            Select Case LCase$(Mid$(args(i), 2))
                Case "dll:activex"
                    dllMode = True: gStrip = False
                Case "dll:standard"
                    dllMode = True: gStrip = True
                Case "dll"
                    Fail "Use /DLL:ACTIVEX or /DLL:STANDARD"
                Case "noicon"
                    gNoIcon = True
                Case Else
                    If LCase$(Mid$(args(i), 2, 4)) = "out:" Then
                        outPath = Mid$(args(i), 6)
                    Else
                        Fail "Unknown switch: " & args(i)
                    End If
            End Select
        Else
            pos(np) = args(i): np = np + 1
        End If
    Next
    If np < 2 Then
        ConPrint "VBStatic [/DLL:ACTIVEX | /DLL:STANDARD] [/NOICON] [/OUT:<file>] <msvbvm60.dll> <main.OBJ> [chainSymbol]", 11
        ConPrint ""
        ConPrint "Fuses the VB runtime into the main OBJ (one object)."
        ConPrint "  /OUT:<file>    save the patched OBJ under this name (default: overwrite <main.OBJ>)"
        ConPrint "  /NOICON        keep Explorer from showing the runtime's icon (runtime icons stay for LoadImage)"
        ConPrint "  (no switch)    EXE"
        ConPrint "  /DLL:ACTIVEX   ActiveX DLL (keeps DllRegisterServer, DllGetClassObject, ...)"
        ConPrint "  /DLL:STANDARD  plain DLL (removes the ActiveX exports; add your own with link /EXPORT:)"
        ConPrint "  chainSymbol    the project's own entry, default ___vbaS (EXE and DLL)"
        mRc = 1
        Exit Sub
    End If
    dllPath = pos(0): objPath = pos(1)
    If np >= 3 Then chain = pos(2)
    If Len(outPath) = 0 Then outPath = objPath

    LoadDll dllPath
    ParseObj objPath
    If FindObjSec(".vbrt") >= 0 Then                  'already patched: harmless, not an error
        AddLog "WARNING: this OBJ is already patched (it contains .vbrt) - left unchanged."
        Exit Sub
    End If
    If Len(chain) = 0 Then chain = "___vbaS"
    If Not HHas(oDefined, chain) Then Fail "The OBJ does not define the entry symbol '" & chain & "' - pass the right one as 3rd argument (chainSymbol)."

    If dllMode Then entryName = "__VBStaticDllEntry@12" Else entryName = "__VBStaticEntry"
    AddLog "entry chains to: " & chain
    PlanResources
    ImportMain
    BuildVbrt chain, entryName, dllMode
    PatchMain
    BuildObjFile res
    WriteBuf outPath, res
    AddLog "written: " & outPath & " (" & res.n & " bytes)"
    AddLog ""
    If dllMode Then
        AddLog "link /DLL /OUT:app.dll /NODEFAULTLIB /ENTRY:_VBStaticDllEntry@12 " & outPath & " <other OBJs>"
        If gStrip Then AddLog "add your own exports, e.g.: /EXPORT:B  (if a name is ambiguous: /EXPORT:B=?B@Module1@@AAGXXZ, see dumpbin /symbols)"
    Else
        AddLog "link /OUT:app.exe /SUBSYSTEM:WINDOWS /NODEFAULTLIB /ENTRY:_VBStaticEntry " & outPath & " <other OBJs>"
    End If
    ConPrint "Done.", 10
    Exit Sub
EH:
    ConPrint "Error: " & Err.Description, 12
    mRc = 1
End Sub

Public Sub Main()
    ConInit
    RunMain
    ConDone
    If Not InIDE() Then ExitProcess mRc                'real exit code for callers (0 = success)
End Sub

Private Function InIDE() As Boolean
    Debug.Assert MarkIDE(InIDE)                        'only executed inside the IDE
End Function

Private Function MarkIDE(ByRef b As Boolean) As Boolean
    b = True
    MarkIDE = True
End Function

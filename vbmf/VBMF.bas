Attribute VB_Name = "modVBMF"
Option Explicit
'==============================================================================
' VBMF  -  embed an application manifest into the main OBJ of a VB6 project
'
' Stand-alone: this single module is the whole tool (no other module / project needed).
'
' Usage (command line):
'     VBMF.exe [/ID:<n>] [/OUT:<file>] <main.OBJ> <main.exe.manifest>
'
'   <main.OBJ>            the project's main object (the one that holds the project's
'                         resources - icon, version info - e.g. Projekt1.OBJ / out1.OBJ).
'   <main.exe.manifest>   the manifest file (XML, embedded unchanged as RT_MANIFEST).
'   /ID:<n>               resource ID of the manifest. Default 1 (EXE). Use /ID:2 for a DLL
'                         (ISOLATIONAWARE_MANIFEST_RESOURCE_ID).
'   /OUT:<file>           write the patched object to <file>. Without /OUT the main OBJ is
'                         overwritten in place (it is built completely in memory first).
'
' The manifest is stored as resource 24 (RT_MANIFEST) / <ID> / language 1033 in the OBJ's
' resource tree (.rsrc$01/.rsrc$02). Everything else in the tree (icons, version info,
' resources merged by VBStatic ...) is kept. A manifest that is already in the OBJ with
' the same ID (any language) is replaced, so running VBMF again with a new manifest is
' harmless. If the OBJ has no resource sections, they are created.
'
' VBMF and VBStatic can be run in any order on the main OBJ:
'     VBMF.exe     Projekt1.OBJ Projekt1.exe.manifest
'     VBStatic.exe msvbvm60.dll Projekt1.OBJ
'     link /OUT:app.exe /SUBSYSTEM:WINDOWS /NODEFAULTLIB /ENTRY:_VBStaticEntry Projekt1.OBJ ...
'
' Output goes to a console (AllocConsole), or into the calling cmd window when the exe
' is switched to the console subsystem (same behaviour as VBStatic.exe).
'==============================================================================

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ExitProcess Lib "Kernel32" (ByVal uExitCode As Long)
Private Declare Function AllocConsole Lib "Kernel32" () As Long
Private Declare Function FreeConsole Lib "Kernel32" () As Long
Private Declare Function GetStdHandle Lib "Kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function SetConsoleTitle Lib "Kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsole Lib "Kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Sub SetConsoleTextAttribute Lib "Kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long)
Private Declare Function SetConsoleMode Lib "Kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function WriteFile Lib "Kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Const RT_MANIFEST As Long = 24
Private Const LANG_EN_US As Long = 1033

Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const R_DIR32NB As Long = 7

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

'---- the OBJ -------------------------------------------------------------------
Private ob() As Byte
Private oNs As Long, oSymp As Long, oNsym As Long, oStOff As Long, oStLen As Long
Private oVal() As Long, oSec() As Long, oCls() As Long

'---- resources -----------------------------------------------------------------
Private rn(1 To 4096) As TRN
Private rnN As Long
Private hLeaf As THash
Private rsr01 As TBuf, rsr02 As TBuf, rsrRel As TBuf
Private rLeaf(1 To 4096) As Long, rLeafN As Long
Private hStr As THash, rCur As Long

'---- COFF writer state ---------------------------------------------------------
Private cs() As TSec
Private csN As Long
Private symTab As TBuf, strTab As TBuf, symN As Long
Private mRc As Long                                 'process exit code: 0 = ok, 1 = error
Private hOut As Long, ownConsole As Boolean, outIsConsole As Boolean

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

Private Sub BPatchW(b As TBuf, ByVal off As Long, ByVal v As Long)
    b.d(off) = CByte(v And &HFF&)
    b.d(off + 1) = CByte((v And &HFF00&) \ &H100&)
End Sub

Private Sub BPatchL(b As TBuf, ByVal off As Long, ByVal v As Long)
    CopyMemory b.d(off), v, 4
End Sub

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
' tiny string -> Long hash table
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
' console output
'==============================================================================
Private Sub ConInit()
    Dim mode As Long
    ownConsole = (AllocConsole() <> 0)             'fails when a console already exists (console subsystem)
    hOut = GetStdHandle(STD_OUTPUT_HANDLE)
    If ownConsole Then SetConsoleTitle "VBMF"
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
    Err.Raise vbObjectError + 1, "VBMF", msg
End Sub

Private Sub AddLog(ByVal s As String)
    If Left$(s, 7) = "WARNING" Then ConPrint s, 14 Else ConPrint s
End Sub

'==============================================================================
' OBJ parsing
'==============================================================================

Private Sub ParseObj(ByVal path As String)
    Dim i As Long, eo As Long, na As Long
    ob = ReadAll(path)
    If UBound(ob) < 19 Then Fail "Not an i386 COFF object: " & path
    If W16(ob, 0) <> &H14C Then Fail "Not an i386 COFF object: " & path
    oNs = W16(ob, 2): oSymp = L32(ob, 8): oNsym = L32(ob, 12)
    oStOff = oSymp + oNsym * 18: oStLen = L32(ob, oStOff)
    ReDim oVal(0 To oNsym - 1): ReDim oSec(0 To oNsym - 1): ReDim oCls(0 To oNsym - 1)
    i = 0
    Do While i < oNsym
        eo = oSymp + i * 18
        oVal(i) = L32(ob, eo + 8)
        oSec(i) = W16(ob, eo + 12): If oSec(i) > 32767 Then oSec(i) = oSec(i) - 65536
        oCls(i) = ob(eo + 16): na = ob(eo + 17)
        i = i + 1 + na
    Loop
    AddLog "OBJ: " & oNs & " sections, " & oNsym & " symbol entries"
End Sub

Private Function FindObjSec(ByVal nm As String) As Long
    Dim i As Long
    For i = 0 To oNs - 1
        If CStrAt(ob, 20 + i * 40, 8) = nm Then FindObjSec = i: Exit Function
    Next
    FindObjSec = -1
End Function

'==============================================================================
' resource tree
'==============================================================================
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
            rn(c).dOff = HGet(hLeaf, CStr(o))
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
        BPutBytes rsr02, ob, rn(c).dOff, rn(c).dSize
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

Private Sub LinkChild(ByVal par As Long, ByVal kid As Long)
    Dim tail As Long
    rn(kid).sib = 0
    If rn(par).child = 0 Then
        rn(par).child = kid
    Else
        tail = rn(par).child
        Do While rn(tail).sib <> 0
            tail = rn(tail).sib
        Loop
        rn(tail).sib = kid
    End If
    rn(par).nChild = rn(par).nChild + 1
End Sub

Private Function NewResDir(ByVal resId As Long) As Long
    Dim z As TRN, n As Long
    n = NewResNode()
    rn(n) = z
    rn(n).isDir = True: rn(n).id = resId
    NewResDir = n
End Function

Private Function FindResById(ByVal par As Long, ByVal resId As Long) As Long
    Dim c As Long
    c = rn(par).child
    Do While c <> 0
        If (Not rn(c).isNamed) And rn(c).id = resId Then FindResById = c: Exit Function
        c = rn(c).sib
    Loop
End Function

'==============================================================================
' COFF writer
'==============================================================================
Private Function AddSym(ByVal nm As String, ByVal value As Long, ByVal sec As Long, ByVal typ As Long, ByVal cls As Long, Optional ByVal aux As Boolean = False) As Long
    Dim i As Long
    AddSym = symN
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
' add / replace one resource
'==============================================================================
'Adds the resource resType/resName/resLang (numeric IDs) with the bytes in data() to the
'loaded OBJ. Every resource with the same type AND name (any language) is replaced.
'Everything else in the OBJ's resource tree (icons, version, ...) is kept. When the OBJ has
'no resource sections yet, .rsrc$01 / .rsrc$02 are created.
Private Sub SetResource(ByVal resType As Long, ByVal resName As Long, ByVal resLang As Long, data() As Byte)
    Dim cnt As Long, dOff As Long, i1 As Long, i2 As Long, h1 As Long, h2 As Long, k As Long
    Dim rs1 As Long, rp1 As Long, rr1 As Long, nr1 As Long, rp2 As Long, va As Long, si As Long
    Dim root As Long, tNode As Long, nNode As Long, lf As Long, c As Long, prev As Long, nx As Long
    Dim old As Long, c1 As Long, c2 As Long, r1 As Long, r2 As Long
    Dim z As TRN
    cnt = UBound(data) - LBound(data) + 1
    If cnt <= 0 Then Fail "Empty resource data"
    rnN = 0
    HInit hLeaf

    'the payload is appended to the OBJ image: LayoutResources copies leaf data out of 'ob'
    dOff = UBound(ob) + 1
    ReDim Preserve ob(0 To dOff + cnt - 1)
    CopyMemory ob(dOff), data(LBound(data)), cnt

    i1 = FindObjSec(".rsrc$01"): i2 = FindObjSec(".rsrc$02")
    If (i1 < 0) <> (i2 < 0) Then Fail "The OBJ has only one of .rsrc$01 / .rsrc$02"
    If i1 >= 0 Then
        h1 = 20 + i1 * 40: h2 = 20 + i2 * 40
        rs1 = L32(ob, h1 + 16): rp1 = L32(ob, h1 + 20): rr1 = L32(ob, h1 + 24): nr1 = W16(ob, h1 + 32)
        rp2 = L32(ob, h2 + 20)
        If rs1 > 0 Then
            For k = 0 To nr1 - 1
                va = L32(ob, rr1 + 10 * k): si = L32(ob, rr1 + 10 * k + 4)
                HSet hLeaf, CStr(va), rp2 + oVal(si) + L32(ob, rp1 + va)
            Next
            root = ParseResDir(ob, rp1, 0, False)
        End If
    End If
    If root = 0 Then
        root = NewResNode()
        rn(root) = z
        rn(root).isDir = True
    End If

    'type directory (created when missing); drop an existing entry with the same name
    tNode = FindResById(root, resType)
    If tNode = 0 Then
        tNode = NewResDir(resType)
        LinkChild root, tNode
    End If
    c = rn(tNode).child
    Do While c <> 0
        nx = rn(c).sib
        If (Not rn(c).isNamed) And rn(c).id = resName Then
            old = old + CountLeaves(c)
            If prev = 0 Then rn(tNode).child = nx Else rn(prev).sib = nx
            rn(tNode).nChild = rn(tNode).nChild - 1
        Else
            prev = c
        End If
        c = nx
    Loop
    If old > 0 Then AddLog "resources: replaced the existing resource " & resType & "/" & resName & " (" & old & " item(s))"

    nNode = NewResDir(resName)
    LinkChild tNode, nNode
    lf = NewResNode()
    rn(lf) = z
    rn(lf).id = resLang: rn(lf).dSize = cnt: rn(lf).cp = 0: rn(lf).dOff = dOff
    LinkChild nNode, lf

    LayoutResources root
    If i1 >= 0 Then
        c1 = i1 + 1: c2 = i2 + 1
        If cs(c2).symIdx < 0 Then Fail "No section symbol for .rsrc$02 in the OBJ"
        cs(c1).data.n = 0: BPutBytes cs(c1).data, rsr01.d, 0, rsr01.n: cs(c1).nrel = 0
        cs(c2).data.n = 0: BPutBytes cs(c2).data, rsr02.d, 0, rsr02.n
        For k = 0 To rsrRel.n \ 4 - 1
            AddRel c1, BGetL(rsrRel, k * 4), cs(c2).symIdx, R_DIR32NB
        Next
        AddLog "resources: OBJ resource sections updated (" & rLeafN & " items, " & cnt & " bytes added as " & resType & "/" & resName & "/" & resLang & ")"
    Else
        r1 = NewSec(".rsrc$01", &H40300040)
        r2 = NewSec(".rsrc$02", &H40300040)
        BPutBytes cs(r1).data, rsr01.d, 0, rsr01.n
        BPutBytes cs(r2).data, rsr02.d, 0, rsr02.n
        For k = 0 To rsrRel.n \ 4 - 1
            AddRel r1, BGetL(rsrRel, k * 4), cs(r2).symIdx, R_DIR32NB
        Next
        AddLog "resources: OBJ had none - resource sections added (" & cnt & " bytes as " & resType & "/" & resName & "/" & resLang & ")"
    End If
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
    Dim objPath As String, mfPath As String, outPath As String, sw As String
    Dim resId As Long, mf() As Byte, txt As String, res As TBuf
    On Error GoTo EH
    resId = 1
    n = SplitArgs(Command$, args)
    For i = 0 To n - 1
        If Left$(args(i), 1) = "/" Then
            sw = LCase$(Mid$(args(i), 2))
            If Left$(sw, 4) = "out:" Then
                outPath = Mid$(args(i), 6)
            ElseIf Left$(sw, 3) = "id:" Then
                resId = CLng(Val(Mid$(args(i), 5)))
                If resId < 1 Or resId > 65535 Then Fail "Bad /ID value (1..65535): " & args(i)
            Else
                Fail "Unknown switch: " & args(i)
            End If
        Else
            pos(np) = args(i): np = np + 1
        End If
    Next
    If np < 2 Then
        ConPrint "VBMF [/ID:<n>] [/OUT:<file>] <main.OBJ> <main.exe.manifest>", 11
        ConPrint ""
        ConPrint "Embeds the manifest into the main OBJ (as RT_MANIFEST resource)."
        ConPrint "  /ID:<n>        resource ID, default 1 (EXE); use 2 for a DLL"
        ConPrint "  /OUT:<file>    save the patched OBJ under this name (default: overwrite <main.OBJ>)"
        mRc = 1
        Exit Sub
    End If
    If np > 2 Then Fail "Too many arguments (expected <main.OBJ> <manifest>)"
    objPath = pos(0): mfPath = pos(1)
    If Len(outPath) = 0 Then outPath = objPath

    mf = ReadAll(mfPath)
    txt = StrConv(mf, vbUnicode)
    If InStr(1, txt, "<assembly", vbTextCompare) = 0 Then
        AddLog "WARNING: " & mfPath & " does not look like an application manifest (no <assembly> element)"
    End If
    AddLog "manifest: " & mfPath & " (" & (UBound(mf) + 1) & " bytes)"

    ParseObj objPath
    ImportMain
    SetResource RT_MANIFEST, resId, LANG_EN_US, mf
    BuildObjFile res
    WriteBuf outPath, res
    AddLog "written: " & outPath & " (" & res.n & " bytes)"
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

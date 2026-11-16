Attribute VB_Name = "modVBExtern"
Option Explicit
'==============================================================================
' VBExtern - turn VB6 "Declare" statements into ordinary linker externs.
'
' VB6 emits one 25-byte thunk per Declare in the project .OBJ:
'
'     a1 <cache>           mov  eax, [cache]
'     0b c0                or   eax, eax
'     74 02                je   +2
'     ff e0                jmp  eax
'     68 <declare_struct>  push offset struct   ; { &"lib", &"func", ... }
'     b8 <DllFunctionCall> mov  eax, offset DllFunctionCall
'     ff d0                call eax
'     ff e0                jmp  eax
'
' For every thunk whose Lib name matches (default "<extern>") this tool writes
'
'     e9 <rel32>           jmp  _<FuncName>   (REL32 reloc to a NEW undefined symbol)
'     cc cc cc ...
'
' so the linker resolves the function against any other .obj/.lib you link in
' (e.g. one compiled from C).  The two old relocations that pointed into the
' removed bytes are left alone: they now only touch the int3 padding behind the
' jmp, which is never executed.
'
' VB6 side:
'   Private Declare Function MyFunction Lib "<extern>" Alias "MyFunction@8" _
'           (ByVal a As Long, ByVal b As Long) As Long
' The extern symbol is  <prefix> & <Alias or declared name>  (prefix = "_").
'
' Usage:  VBExtern Projekt1.OBJ                      patch in place
'         VBExtern Projekt1.OBJ /OUT:out.obj          write somewhere else
'         VBExtern Projekt1.OBJ /LIST                 only list the Declares
'         VBExtern Projekt1.OBJ /PLACEHOLDER:kernel32 use another Lib name as marker
'         VBExtern Projekt1.OBJ /PREFIX:              no symbol prefix
'
' Build: Standard EXE, remove Form1, add this module, Startup Object = Sub Main,
' and add /SUBSYSTEM:CONSOLE to the linker switches (see VBExtern.vbp).
'==============================================================================

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private Const SCN_CNT_CODE As Long = &H20&
Private Const SCN_LNK_NRELOC_OVFL As Long = &H1000000
Private Const REL_DIR32 As Long = 6
Private Const REL_REL32 As Long = &H14&
Private Const SYM_EXTERNAL As Long = 2
Private Const THUNK_LEN As Long = 25
Private Const OFF_CACHE As Long = 1
Private Const OFF_STRUCT As Long = 12
Private Const OFF_CALLEE As Long = 17

' ---- object file model ------------------------------------------------------
Private F() As Byte            ' file bytes up to the start of the symbol table
Private SymBuf() As Byte       ' symbol table (18 bytes per record, aux included)
Private StrBuf() As Byte       ' string table including its 4-byte length field
Private mSecCount As Long
Private mSymOff As Long
Private mSymCount As Long

' ---- declares found ---------------------------------------------------------
Private nDecl As Long
Private dSec() As Long
Private dOff() As Long
Private dLib() As String
Private dFunc() As String
Private dThunk() As String

' ---- console ----------------------------------------------------------
Private mStdOut As Long

'------------------------------------------------------------------------------
Public Sub Main()
    Dim rc As Long
    On Error GoTo EH
    mStdOut = GetStdHandle(-11&)           ' STD_OUTPUT_HANDLE (project is linked /SUBSYSTEM:CONSOLE)
    If mStdOut = -1& Then mStdOut = 0
    rc = RunTool(Command$)
    Finish rc
    Exit Sub
EH:
    Out "error: " & Err.Description
    Finish 1
End Sub

Private Function RunTool(ByVal cmdLine As String) As Long
    Dim args As Collection, i As Long, a As String
    Dim inPath As String, outPath As String, prefix As String
    Dim libs() As String, nLibs As Long, doList As Boolean
    Dim k As Long, todo As Long, target As String
    Dim nm As String, p As Long, sOpt As String, sVal As String

    prefix = "_"
    Set args = SplitArgs(cmdLine)
    i = 1
    Do While i <= args.Count
        a = args(i)
        If Left$(a, 1) = "/" Then
            p = InStr(a, ":")
            If p > 0 Then
                sOpt = UCase$(Mid$(a, 2, p - 2))
                sVal = Mid$(a, p + 1)
            Else
                sOpt = UCase$(Mid$(a, 2))
                sVal = ""
            End If
            Select Case sOpt
                Case "OUT"
                    If Len(sVal) = 0 Then Err.Raise vbObjectError + 1, , "missing file name in " & a
                    outPath = sVal
                Case "PLACEHOLDER"
                    If Len(sVal) = 0 Then Err.Raise vbObjectError + 1, , "missing library name in " & a
                    ReDim Preserve libs(nLibs): libs(nLibs) = LCase$(sVal): nLibs = nLibs + 1
                Case "PREFIX"
                    prefix = sVal
                Case "LIST"
                    doList = True
                Case "?", "HELP"
                    Call Usage: RunTool = 0: Exit Function
                Case Else
                    Err.Raise vbObjectError + 1, , "unknown option: " & a
            End Select
        ElseIf Len(inPath) = 0 Then
            inPath = a
        Else
            Err.Raise vbObjectError + 1, , "unexpected argument: " & a
        End If
        i = i + 1
    Loop
    If Len(inPath) = 0 Then Call Usage: RunTool = 1: Exit Function
    If nLibs = 0 Then ReDim libs(0): libs(0) = "<extern>": nLibs = 1

    LoadObj inPath
    FindDeclares
    If nDecl = 0 Then
        Out "No VB6 Declare thunks found (already patched, or no Declare statements)."
        RunTool = 0: Exit Function
    End If

    Out "Declares found in " & inPath & ":"
    For k = 0 To nDecl - 1
        Out "  [" & IIf(IsMarker(dLib(k), libs, nLibs), "EXTERN", "      ") & "] Declare " & dFunc(k) & _
            " Lib """ & dLib(k) & """  thunk " & SecName(dSec(k)) & "+0x" & Hex$(dOff(k)) & _
            IIf(Len(dThunk(k)) > 0, " (" & dThunk(k) & ")", "")
    Next
    If doList Then RunTool = 0: Exit Function

    For k = 0 To nDecl - 1
        If IsMarker(dLib(k), libs, nLibs) Then
            todo = todo + 1
            If Len(dFunc(k)) = 0 Or Left$(dFunc(k), 1) = "#" Then
                Err.Raise vbObjectError + 1, , "ordinal / empty function names cannot be turned into a symbol: '" & dFunc(k) & "'"
            End If
        End If
    Next
    If todo = 0 Then
        Out "Nothing to patch: no Declare uses the selected Lib name(s)."
        RunTool = 0: Exit Function
    End If

    Out ""
    For k = 0 To nDecl - 1
        If IsMarker(dLib(k), libs, nLibs) Then
            nm = PatchOne(k, prefix)
            Out "  " & dFunc(k) & "  ->  undefined external symbol " & nm
        End If
    Next

    If Len(outPath) > 0 Then
        target = outPath
    Else
        target = inPath
    End If
    SaveObj target
    Out "Patched " & todo & " declare(s) -> " & target
    Out "Link the result together with an object/library that defines those symbols."
    RunTool = 0
End Function

Private Sub Usage()
    Out "VBExtern <file.OBJ> [/OUT:file] [/LIST] [/PLACEHOLDER:name]... [/PREFIX:text]"
    Out "  Turns Declare ... Lib ""<extern>"" thunks into undefined linker externs."
    Out "  /OUT:file          write the patched object here (default: patch in place)"
    Out "  /LIST              only list the Declares found"
    Out "  /PLACEHOLDER:name  Lib name that marks an extern declare (default ""<extern>"")"
    Out "  /PREFIX:text       symbol prefix (default ""_"")"
End Sub

Private Function IsMarker(ByVal libName As String, ByRef libs() As String, ByVal nLibs As Long) As Boolean
    Dim j As Long
    For j = 0 To nLibs - 1
        If LCase$(libName) = libs(j) Then IsMarker = True: Exit Function
    Next
End Function

'------------------------------------------------------------------------------
' Loading / saving
'------------------------------------------------------------------------------
Private Sub LoadObj(ByVal path As String)
    Dim h As Long, n As Long, i As Long, st As Long, slen As Long, hd As Long

    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 1, , "file not found: " & path
    n = FileLen(path)
    If n < 20 Then Err.Raise vbObjectError + 1, , "file too small to be a COFF object"
    ReDim F(n - 1)
    h = FreeFile
    Open path For Binary Access Read As #h
    Get #h, , F
    Close #h

    If U16(F, 0) <> &H14C& Then Err.Raise vbObjectError + 1, , "not an x86 (i386) COFF object"
    If U16(F, 16) <> 0 Then Err.Raise vbObjectError + 1, , "object has an optional header; not a plain .obj"
    mSecCount = U16(F, 2)
    mSymOff = I32(F, 8)
    mSymCount = I32(F, 12)

    st = mSymOff + 18& * mSymCount
    If mSymOff < 20 + 40& * mSecCount Or st + 4 > n Then Err.Raise vbObjectError + 1, , "unsupported symbol table layout"
    slen = I32(F, st)
    If slen < 4 Or st + slen <> n Then Err.Raise vbObjectError + 1, , "unsupported layout: string table is not at end of file"

    For i = 0 To mSecCount - 1
        hd = 20 + 40& * i
        If (I32(F, hd + 36) And SCN_LNK_NRELOC_OVFL) <> 0 Then Err.Raise vbObjectError + 1, , "extended relocation counts are not supported"
        If I32(F, hd + 28) <> 0 Or U16(F, hd + 34) <> 0 Then Err.Raise vbObjectError + 1, , "COFF line numbers are not supported"
    Next

    If mSymCount > 0 Then
        ReDim SymBuf(18& * mSymCount - 1)
        For i = 0 To UBound(SymBuf): SymBuf(i) = F(mSymOff + i): Next
    Else
        ReDim SymBuf(0)
    End If
    ReDim StrBuf(slen - 1)
    For i = 0 To slen - 1: StrBuf(i) = F(st + i): Next
End Sub

Private Sub SaveObj(ByVal path As String)
    Dim h As Long
    PutI32 F, 12, mSymCount                ' new symbol count in the header
    ReDim Preserve F(mSymOff - 1)          ' everything in front of the symbol table
    If Len(Dir$(path)) > 0 Then Kill path
    h = FreeFile
    Open path For Binary Access Write As #h
    Put #h, , F
    Put #h, , SymBuf
    Put #h, , StrBuf
    Close #h
End Sub

'------------------------------------------------------------------------------
' Byte helpers (little endian)
'------------------------------------------------------------------------------
Private Function U16(ByRef b() As Byte, ByVal p As Long) As Long
    U16 = CLng(b(p)) + CLng(b(p + 1)) * 256&
End Function

Private Function I16(ByRef b() As Byte, ByVal p As Long) As Long
    Dim v As Long
    v = U16(b, p)
    If v >= 32768 Then v = v - 65536
    I16 = v
End Function

Private Function I32(ByRef b() As Byte, ByVal p As Long) As Long
    Dim v As Long
    v = CLng(b(p)) + CLng(b(p + 1)) * 256& + CLng(b(p + 2)) * 65536
    If b(p + 3) >= 128 Then
        v = v + (CLng(b(p + 3)) - 256&) * 16777216
    Else
        v = v + CLng(b(p + 3)) * 16777216
    End If
    I32 = v
End Function

Private Sub PutI32(ByRef b() As Byte, ByVal p As Long, ByVal v As Long)
    b(p) = v And &HFF&
    b(p + 1) = (v And &HFF00&) \ &H100&
    b(p + 2) = (v And &HFF0000) \ &H10000
    b(p + 3) = ((v And &HFF000000) \ &H1000000) And &HFF&
End Sub

Private Sub PutU16(ByRef b() As Byte, ByVal p As Long, ByVal v As Long)
    b(p) = v And &HFF&
    b(p + 1) = (v \ &H100&) And &HFF&
End Sub

'------------------------------------------------------------------------------
' Sections / symbols
'------------------------------------------------------------------------------
Private Function SecRaw(ByVal si As Long) As Long
    SecRaw = I32(F, 20 + 40& * si + 20)
End Function

Private Function SecSize(ByVal si As Long) As Long
    SecSize = I32(F, 20 + 40& * si + 16)
End Function

Private Function SecName(ByVal si As Long) As String
    Dim k As Long, s As String
    For k = 0 To 7
        If F(20 + 40& * si + k) = 0 Then Exit For
        s = s & Chr$(F(20 + 40& * si + k))
    Next
    SecName = s
End Function

Private Function ReadZ(ByRef b() As Byte, ByVal p As Long) As String
    Dim s As String
    Do While p <= UBound(b)
        If b(p) = 0 Then Exit Do
        s = s & Chr$(b(p))
        p = p + 1
    Loop
    ReadZ = s
End Function

Private Function SymName(ByVal i As Long) As String
    Dim p As Long, k As Long, s As String
    p = i * 18&
    If I32(SymBuf, p) = 0 Then
        SymName = ReadZ(StrBuf, I32(SymBuf, p + 4))     ' offsets include the 4-byte length field
    Else
        For k = 0 To 7
            If SymBuf(p + k) = 0 Then Exit For
            s = s & Chr$(SymBuf(p + k))
        Next
        SymName = s
    End If
End Function

' Index of an EXTERNAL symbol called nm; appends a new undefined one if absent.
Private Function AddExtern(ByVal nm As String) As Long
    Dim i As Long, p As Long, k As Long, off As Long, n As Long

    i = 0
    Do While i < mSymCount
        If SymBuf(i * 18& + 16) = SYM_EXTERNAL Then
            If SymName(i) = nm Then AddExtern = i: Exit Function
        End If
        i = i + 1 + SymBuf(i * 18& + 17)
    Loop

    p = mSymCount * 18&
    ReDim Preserve SymBuf(p + 17)
    For k = 0 To 17: SymBuf(p + k) = 0: Next
    If Len(nm) <= 8 Then
        For k = 1 To Len(nm): SymBuf(p + k - 1) = Asc(Mid$(nm, k, 1)): Next
    Else
        off = UBound(StrBuf) + 1
        n = Len(nm) + 1
        ReDim Preserve StrBuf(off + n - 1)
        For k = 1 To Len(nm): StrBuf(off + k - 1) = Asc(Mid$(nm, k, 1)): Next
        StrBuf(off + n - 1) = 0
        PutI32 StrBuf, 0, off + n                       ' new string table length
        PutI32 SymBuf, p + 4, off                       ' name = (0, strtab offset)
    End If
    SymBuf(p + 14) = &H20                               ' type: function
    SymBuf(p + 16) = SYM_EXTERNAL                       ' storage class
    AddExtern = mSymCount
    mSymCount = mSymCount + 1
End Function

'------------------------------------------------------------------------------
' Relocations
'------------------------------------------------------------------------------
' Position (in F) of the relocation record at `site` in section si, or -1.
Private Function FindReloc(ByVal si As Long, ByVal site As Long) As Long
    Dim hd As Long, relPtr As Long, nRel As Long, r As Long, q As Long
    hd = 20 + 40& * si
    relPtr = I32(F, hd + 24)
    nRel = U16(F, hd + 32)
    FindReloc = -1
    For r = 0 To nRel - 1
        q = relPtr + r * 10&
        If I32(F, q) = site Then FindReloc = q: Exit Function
    Next
End Function

Private Function IsDir32(ByVal si As Long, ByVal site As Long) As Boolean
    Dim q As Long
    q = FindReloc(si, site)
    If q >= 0 Then IsDir32 = (U16(F, q + 8) = REL_DIR32)
End Function

' Follow the DIR32 relocation at `site`: target section index + offset.
Private Function Resolve(ByVal si As Long, ByVal site As Long, ByRef outSec As Long, ByRef outOff As Long) As Boolean
    Dim q As Long, idx As Long, secNum As Long, value As Long, addend As Long

    If SecRaw(si) = 0 Or site < 0 Or site + 4 > SecSize(si) Then Exit Function
    q = FindReloc(si, site)
    If q < 0 Then Exit Function
    If U16(F, q + 8) <> REL_DIR32 Then Exit Function
    idx = I32(F, q + 4)
    If idx < 0 Or idx >= mSymCount Then Exit Function
    secNum = I16(SymBuf, idx * 18& + 12)
    If secNum <= 0 Or secNum > mSecCount Then Exit Function
    value = I32(SymBuf, idx * 18& + 8)
    addend = I32(F, SecRaw(si) + site)
    outSec = secNum - 1
    outOff = value + addend
    Resolve = True
End Function

Private Function ReadZAt(ByVal si As Long, ByVal off As Long, ByRef s As String) As Boolean
    If SecRaw(si) = 0 Or off < 0 Or off >= SecSize(si) Then Exit Function
    s = ReadZ(F, SecRaw(si) + off)
    ReadZAt = True
End Function

'------------------------------------------------------------------------------
' Finding and patching the Declare thunks
'------------------------------------------------------------------------------
Private Function IsThunk(ByVal p As Long) As Boolean
    If F(p) <> &HA1 Then Exit Function
    If F(p + 5) <> &HB Or F(p + 6) <> &HC0 Then Exit Function
    If F(p + 7) <> &H74 Or F(p + 8) <> 2 Then Exit Function
    If F(p + 9) <> &HFF Or F(p + 10) <> &HE0 Then Exit Function
    If F(p + 11) <> &H68 Then Exit Function
    If F(p + 16) <> &HB8 Then Exit Function
    If F(p + 21) <> &HFF Or F(p + 22) <> &HD0 Then Exit Function
    If F(p + 23) <> &HFF Or F(p + 24) <> &HE0 Then Exit Function
    IsThunk = True
End Function

Private Sub AddDecl(ByVal si As Long, ByVal st As Long, ByVal libName As String, ByVal func As String)
    Dim i As Long
    ReDim Preserve dSec(nDecl): ReDim Preserve dOff(nDecl)
    ReDim Preserve dLib(nDecl): ReDim Preserve dFunc(nDecl): ReDim Preserve dThunk(nDecl)
    dSec(nDecl) = si: dOff(nDecl) = st: dLib(nDecl) = libName: dFunc(nDecl) = func
    ' friendly name of the thunk symbol, if there is one
    i = 0
    Do While i < mSymCount
        If SymBuf(i * 18& + 16) = SYM_EXTERNAL And I16(SymBuf, i * 18& + 12) = si + 1 And I32(SymBuf, i * 18& + 8) = st Then
            dThunk(nDecl) = SymName(i): Exit Do
        End If
        i = i + 1 + SymBuf(i * 18& + 17)
    Loop
    nDecl = nDecl + 1
End Sub

Private Sub FindDeclares()
    Dim si As Long, hd As Long, flags As Long, raw As Long, size As Long, st As Long
    Dim sSec As Long, sOff As Long, lSec As Long, lOff As Long, fSec As Long, fOff As Long
    Dim libName As String, func As String

    nDecl = 0
    For si = 0 To mSecCount - 1
        hd = 20 + 40& * si
        flags = I32(F, hd + 36)
        raw = I32(F, hd + 20)
        size = I32(F, hd + 16)
        If (flags And SCN_CNT_CODE) <> 0 And raw <> 0 Then
            For st = 0 To size - THUNK_LEN
                If IsThunk(raw + st) Then
                  If IsDir32(si, st + OFF_CACHE) And IsDir32(si, st + OFF_STRUCT) And IsDir32(si, st + OFF_CALLEE) Then
                    If Resolve(si, st + OFF_STRUCT, sSec, sOff) Then
                      If Resolve(sSec, sOff, lSec, lOff) Then
                        If Resolve(sSec, sOff + 4, fSec, fOff) Then
                          If ReadZAt(lSec, lOff, libName) Then
                            If ReadZAt(fSec, fOff, func) Then AddDecl si, st, libName, func
                          End If
                        End If
                      End If
                    End If
                  End If
                End If
            Next
        End If
    Next
End Sub

' Rewrites declare k into  jmp <extern>  and returns the extern symbol name.
'
' Why the two other relocations of the old thunk are NOT removed:
'   The old thunk had three relocated imm32 fields: the cache variable (+1),
'   the Declare struct (+12) and DllFunctionCall (+17).  Only the first one is
'   reused: it becomes the REL32 of the new jmp.  The other two now point into
'   the int3 padding behind the jmp, and they are deliberately left in the
'   relocation table because
'   - deleting a record shrinks the section's relocation table, so everything
'     behind it (later relocation tables, symbol table, string table) would
'     have to be moved and every PointerToRelocations / PointerToSymbolTable
'     field fixed up.  That is a lot of error-prone byte shuffling in VB6 for
'     no functional gain.
'   - they are harmless: the linker only writes two addresses into padding that
'     is never executed (the jmp is unconditional), and the symbols they name
'     (the struct and the DllFunctionCall stub) still exist in this object, so
'     nothing becomes unresolved.
'   Cost: two stray addresses in the padding, two extra base-relocation entries
'   when linking a DLL, and the unused struct/stub are not discarded.
Private Function PatchOne(ByVal k As Long, ByVal prefix As String) As String
    Dim si As Long, st As Long, bp As Long, idx As Long, q As Long, j As Long, nm As String

    si = dSec(k): st = dOff(k)
    bp = SecRaw(si) + st
    nm = prefix & dFunc(k)
    idx = AddExtern(nm)

    F(bp) = &HE9                                     ' jmp rel32
    For j = 1 To 4: F(bp + j) = 0: Next
    For j = 5 To THUNK_LEN - 1: F(bp + j) = &HCC: Next

    q = FindReloc(si, st + OFF_CACHE)                  ' the old "mov eax,[cache]" imm32 ...
    PutI32 F, q + 4, idx                               ' ... becomes REL32 against the extern
    PutU16 F, q + 8, REL_REL32
    PatchOne = nm
End Function

'------------------------------------------------------------------------------
' Command line / output
'------------------------------------------------------------------------------
Private Function SplitArgs(ByVal s As String) As Collection
    Dim c As New Collection, i As Long, ch As String, cur As String
    Dim inQ As Boolean, started As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inQ = Not inQ: started = True
        ElseIf (ch = " " Or ch = vbTab) And Not inQ Then
            If started Then c.Add cur: cur = "": started = False
        Else
            cur = cur & ch: started = True
        End If
    Next
    If started Then c.Add cur
    Set SplitArgs = c
End Function

Private Sub Out(ByVal s As String)
    Dim b() As Byte, w As Long
    s = s & vbCrLf
    Debug.Print s;                                     ' IDE: goes to the Immediate window
    If mStdOut <> 0 Then
        b = StrConv(s, vbFromUnicode)
        WriteFile mStdOut, b(0), UBound(b) + 1, w, 0&
    End If
End Sub

Private Sub Finish(ByVal rc As Long)
    If Not InIDE() Then ExitProcess rc
End Sub

Private Function InIDE() As Boolean
    Debug.Assert MarkIDE(InIDE)                        ' only executed inside the IDE
End Function

Private Function MarkIDE(ByRef b As Boolean) As Boolean
    b = True
    MarkIDE = True
End Function

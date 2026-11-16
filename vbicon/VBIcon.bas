Attribute VB_Name = "modVBIcon"
Option Explicit
'==============================================================================
' VBIcon - replace the default icon VB6 puts into the project .OBJ.
'
' VB6 emits its own application icon into the .OBJ resource sections as
'
'     RT_GROUP_ICON  #1            (GRPICONDIR with 3 entries)
'     RT_ICON        #30001..30003 (32x32 1bpp, 32x32 4bpp, 16x16 4bpp)
'
' This tool throws exactly those away and puts a new icon set (any number of
' resolutions / colour depths) built from a JPEG / PNG / BMP / ICO file into
' the same slot (group #1).  Everything else is copied bit for bit:
'
'   - your own icons  (e.g. the named group "APPICON" and its RT_ICON data)
'   - version info, manifests, bitmaps, strings, RCDATA ... every other resource
'   - code, data, imports, relocations, symbols  (only the two .rsrc sections are
'     rebuilt, the rest of the file is re-laid-out unchanged)
'
' An RT_ICON that is still referenced by any other (kept) icon group is never
' touched, even if the VB6 group points to it.
'
' How the .OBJ is changed
'   .rsrc$01  resource directory tree + data entries (re-serialised)
'   .rsrc$02  resource data (kept data copied, new icon data appended)
'   Every data entry gets a DIR32NB relocation against the .rsrc$02 section
'   symbol with the data offset stored as addend (same thing a resource compiler
'   would emit).  Symbol indices never change, so no other relocation breaks.
'
' Images
'   GDI+  decodes JPEG/PNG/BMP (and PNG-compressed ICO frames), resamples with
'         the high quality bicubic filter (premultiplied alpha, edge-safe), and
'         encodes the 256x256 PNG frame.
'   GDI   decodes DIB based ICO frames (CreateIconFromResourceEx) and supplies
'         the halftone palette used as fall back for 8 bit.
'   Colour depth reduction uses Atkinson error diffusion, transparency is stored
'   as an AND mask.  Nothing is written to temporary files, all work is done in
'   memory and the output is written once, at the very end.
'
' Usage:  VBIcon <main.obj> <icon.png|jpg|jpeg|bmp|ico> [flags]
'         VBIcon <main.obj> /REMOVE [flags]
'         VBIcon <main.obj> /LIST
'
'   /REMOVE          delete the VB6 icon (group #1 and its 3 images), add nothing
'   /OUT:file.obj    write here instead of overwriting the input (default: overwrite)
'   /RES:256,48,32   sizes to create (also accepts "(256,48,32)").  Default: every
'                    size of 256,128,64,48,32,16 that is <= the smaller side of
'                    the source image
'   /DEPTH:32,8,4    colour depths 1,4,8,24,32 (several allowed, default 32)
'                    256x256 at 32 bit is stored PNG compressed (Vista style)
'   /STRETCH         do not keep the aspect ratio (default: fit, transparent bars)
'   /LIST            only list the resources of the object (VB6 icon is marked *)
'
' Build: Standard EXE, remove Form1, add this module, Startup Object = Sub Main,
' and add /SUBSYSTEM:CONSOLE to the linker switches (see VBIcon.vbp).
'==============================================================================

' ---- kernel32 / ole32 -------------------------------------------------------
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As Long, ByRef phglobal As Long) As Long

' ---- user32 / gdi32 ---------------------------------------------------------
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPaletteEntries Lib "gdi32" (ByVal hPalette As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, ByRef lpPaletteEntries As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' ---- GDI+ flat API ----------------------------------------------------------
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type GpRect
    X As Long
    Y As Long
    W As Long
    H As Long
End Type
Private Type GpBitmapData
    Width As Long
    Height As Long
    Stride As Long
    PixelFormat As Long
    Scan0 As Long
    Reserved As Long
End Type
Private Type GpGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal pStream As Long, ByRef pBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, ByRef pBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nStride As Long, ByVal nFormat As Long, ByVal pScan0 As Long, ByRef pBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal pImage As Long, ByRef nWidth As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal pImage As Long, ByRef nHeight As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal pImage As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal pImage As Long, ByVal nRotateFlipType As Long) As Long
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal pImage As Long, ByVal nPropId As Long, ByRef nSize As Long) As Long
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal pImage As Long, ByVal nPropId As Long, ByVal nSize As Long, ByRef pBuffer As Byte) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal pImage As Long, ByRef pGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal pGraphics As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal pGraphics As Long, ByVal nMode As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal pGraphics As Long, ByVal nMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal pGraphics As Long, ByVal nMode As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal pGraphics As Long, ByVal nArgb As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef pAttr As Long) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal pAttr As Long, ByVal nWrap As Long, ByVal nArgb As Long, ByVal bClamp As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal pAttr As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal pGraphics As Long, ByVal pImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstW As Long, ByVal dstH As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcW As Long, ByVal srcH As Long, ByVal nSrcUnit As Long, ByVal pAttr As Long, ByVal pCallback As Long, ByVal pCallbackData As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal pBitmap As Long, ByRef rc As GpRect, ByVal nFlags As Long, ByVal nFormat As Long, ByRef bd As GpBitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal pBitmap As Long, ByRef bd As GpBitmapData) As Long
Private Declare Function GdipInitializePalette Lib "gdiplus" (ByRef pPalette As Byte, ByVal nPaletteType As Long, ByVal nOptimalColors As Long, ByVal bUseTransparent As Long, ByVal pBitmap As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal pImage As Long, ByVal pStream As Long, ByRef clsidEncoder As GpGuid, ByVal pEncoderParams As Long) As Long

Private Const PF_32ARGB As Long = &H26200A
Private Const PF_32PARGB As Long = &HE200B

' ---- resource constants -----------------------------------------------------
Private Const SCN_LNK_NRELOC_OVFL As Long = &H1000000
Private Const SYM_STATIC As Long = 3
Private Const REL_DIR32NB As Long = 7
Private Const RT_ICON As Long = 3
Private Const RT_GROUP_ICON As Long = 14
Private Const VB6_GROUP_ID As Long = 1
Private Const VB6_ICON_BASE As Long = 30001

' ---- object file model ------------------------------------------------------
Private F() As Byte            ' file bytes up to the start of the symbol table (original layout)
Private SymBuf() As Byte       ' symbol table (18 bytes per record, aux included)
Private StrBuf() As Byte       ' string table including its 4-byte length field
Private mSecCount As Long
Private mSymOff As Long
Private mSymCount As Long
Private mSecR1 As Long         ' index of .rsrc$01
Private mSecR2 As Long         ' index of .rsrc$02
Private mDirHdr(0 To 11) As Byte   ' directory header template (characteristics, time stamp, version)

' ---- new resource sections, filled by BuildRsrc -----------------------------
Private R1() As Byte
Private R1Len As Long
Private R2() As Byte
Private R2Len As Long
Private RR() As Byte           ' relocations of .rsrc$01
Private RRCount As Long

' ---- resource leaves (type / name / language) -------------------------------
Private nRes As Long
Private rTypN() As String      ' type name, "" = numeric id
Private rTypI() As Long
Private rNamN() As String      ' resource name, "" = numeric id
Private rNamI() As Long
Private rLang() As Long
Private rCP() As Long
Private rRsv() As Long
Private rPool() As Long        ' offset of the data in Pool
Private rSize() As Long
Private rSym() As Long         ' symbol the old data entry pointed at (-1 = new)
Private rDel() As Boolean      ' leaf is dropped
Private rVB6() As Boolean      ' leaf belongs to the VB6 default icon
Private Pool() As Byte
Private PoolLen As Long
Private PoolCap As Long

' ---- VB6 icon facts ---------------------------------------------------------
Private vbGrpLang As Long
Private vbIcoLang As Long
Private vbCP As Long
Private vbFound As Boolean

' ---- source image -----------------------------------------------------------
Private mToken As Long
Private mSrc As Long           ' premultiplied 32 bit GDI+ bitmap, the "master"
Private mSrcW As Long
Private mSrcH As Long
Private mStretch As Boolean
Private mPx() As Byte            ' last rendered size as straight BGRA, shared by all depths of that size
Private mPxS As Long
Private BitMask(0 To 7) As Long

' ---- console ----------------------------------------------------------------
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
    Dim objPath As String, icoPath As String, outPath As String, target As String
    Dim doRemove As Boolean, doList As Boolean
    Dim sizes() As Long, nSizes As Long, depths() As Long, nDepths As Long
    Dim p As Long, sOpt As String, sVal As String, nDel As Long, k As Long
    Dim m As Long, std As Variant

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
                Case "REMOVE"
                    doRemove = True
                Case "LIST"
                    doList = True
                Case "STRETCH"
                    mStretch = True
                Case "OUT"
                    If Len(sVal) = 0 Then Err.Raise vbObjectError + 1, , "missing file name in " & a
                    outPath = sVal
                Case "RES"
                    ParseIntList sVal, sizes, nSizes, 8, 256, "/RES"
                Case "DEPTH"
                    ParseIntList sVal, depths, nDepths, 1, 32, "/DEPTH"
                Case "?", "HELP"
                    Call Usage: RunTool = 0: Exit Function
                Case Else
                    Err.Raise vbObjectError + 1, , "unknown option: " & a
            End Select
        ElseIf Len(objPath) = 0 Then
            objPath = a
        ElseIf Len(icoPath) = 0 Then
            icoPath = a
        Else
            Err.Raise vbObjectError + 1, , "unexpected argument: " & a
        End If
        i = i + 1
    Loop
    If Len(objPath) = 0 Then Call Usage: RunTool = 1: Exit Function
    If Not doRemove And Not doList And Len(icoPath) = 0 Then
        Err.Raise vbObjectError + 1, , "no icon file given (use /REMOVE to only delete the VB6 icon)"
    End If
    If doRemove And Len(icoPath) > 0 Then
        Err.Raise vbObjectError + 1, , "/REMOVE and an icon file cannot be combined"
    End If
    For k = 0 To nDepths - 1
        Select Case depths(k)
            Case 1, 4, 8, 24, 32
            Case Else: Err.Raise vbObjectError + 1, , "/DEPTH must be one of 1, 4, 8, 24, 32 (not " & depths(k) & ")"
        End Select
    Next
    If nDepths = 0 Then ReDim depths(0): depths(0) = 32: nDepths = 1
    BitMask(0) = &H80: BitMask(1) = &H40: BitMask(2) = &H20: BitMask(3) = &H10
    BitMask(4) = 8: BitMask(5) = 4: BitMask(6) = 2: BitMask(7) = 1

    LoadObj objPath
    FindRsrcSections
    ParseRsrc
    FindVB6Icon

    If doList Then
        ListResources objPath
        RunTool = 0: Exit Function
    End If

    If doRemove Then
        If Not vbFound Then
            Out "No VB6 default icon (group #" & VB6_GROUP_ID & ") found in " & objPath & " - nothing to remove."
            RunTool = 0: Exit Function
        End If
    Else
        StartGdip
        LoadSource icoPath
        Out "Source: " & icoPath & "  (" & mSrcW & "x" & mSrcH & ")"
        ' sizes: default = every standard size not larger than the source
        If nSizes = 0 Then
            m = mSrcW: If mSrcH < m Then m = mSrcH
            std = Array(256, 128, 64, 48, 32, 16)
            For k = 0 To 5
                If std(k) <= m Then ReDim Preserve sizes(nSizes): sizes(nSizes) = std(k): nSizes = nSizes + 1
            Next
            If nSizes = 0 Then ReDim sizes(0): sizes(0) = 16: nSizes = 1
        End If
        SortDesc sizes, nSizes
        SortDesc depths, nDepths
    End If

    ' ---- drop the VB6 icon ----
    For k = 0 To nRes - 1
        If rVB6(k) Then rDel(k) = True: nDel = nDel + 1
    Next
    If vbFound Then
        Out "Removed VB6 default icon: group #" & VB6_GROUP_ID & " and " & (nDel - CountGroups()) & " image(s)."
    Else
        Out "No VB6 default icon found, the new icon is simply added as group #" & VB6_GROUP_ID & "."
    End If

    ' ---- add the new icon ----
    If Not doRemove Then BuildIcons sizes, nSizes, depths, nDepths

    BuildRsrc
    If Len(outPath) > 0 Then
        target = outPath
    Else
        target = objPath
    End If
    SaveObj target
    Out "Saved " & target & "  (.rsrc$01 " & R1Len & " bytes, .rsrc$02 " & R2Len & " bytes, " & CountLive() & " resources)"
    If Not doRemove Then Out "All resources other than the VB6 default icon were left untouched."
    RunTool = 0
End Function

Private Sub Usage()
    Out "VBIcon <main.obj> <icon.png|jpg|jpeg|bmp|ico> [flags]"
    Out "VBIcon <main.obj> /REMOVE [flags]"
    Out "  Replaces the default icon VB6 stores in the .OBJ (group #1 + its 3 images)."
    Out "  Every other resource (APPICON, version info, ...) is kept bit for bit."
    Out "  /REMOVE            delete the VB6 icon, add nothing"
    Out "  /OUT:file.obj      write here instead of overwriting the input"
    Out "  /RES:256,48,32,16  sizes (default: all of 256,128,64,48,32,16 <= source size)"
    Out "  /DEPTH:32,8,4      colour depths 1,4,8,24,32 (default 32; Atkinson dithering)"
    Out "  /STRETCH           ignore the aspect ratio (default: fit with transparent bars)"
    Out "  /LIST              list the resources of the object, VB6 icon marked with *"
End Sub

Private Sub ParseIntList(ByVal s As String, ByRef arr() As Long, ByRef n As Long, ByVal lo As Long, ByVal hi As Long, ByVal what As String)
    Dim i As Long, ch As String, cur As String, v As Long, j As Long, dup As Boolean
    s = s & ","
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "(", ")", " "
            Case ",", ";", "/", "+"
                If Len(cur) > 0 Then
                    If Not IsNumeric(cur) Then Err.Raise vbObjectError + 1, , "bad number in " & what & ": " & cur
                    v = CLng(cur)
                    If v < lo Or v > hi Then Err.Raise vbObjectError + 1, , what & " value out of range (" & lo & ".." & hi & "): " & v
                    dup = False
                    For j = 0 To n - 1
                        If arr(j) = v Then dup = True
                    Next
                    If Not dup Then ReDim Preserve arr(n): arr(n) = v: n = n + 1
                    cur = ""
                End If
            Case Else
                cur = cur & ch
        End Select
    Next
End Sub

Private Sub SortDesc(ByRef arr() As Long, ByVal n As Long)
    Dim i As Long, j As Long, t As Long
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If arr(j) > arr(i) Then t = arr(i): arr(i) = arr(j): arr(j) = t
        Next
    Next
End Sub

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
        CopyMemory SymBuf(0), F(mSymOff), 18& * mSymCount
    Else
        ReDim SymBuf(0)
    End If
    ReDim StrBuf(slen - 1)
    CopyMemory StrBuf(0), F(st), slen
End Sub

' Writes the complete new file: header, section headers, raw data (4 byte
' aligned like the compiler does it), relocations, symbols, strings.
Private Sub SaveObj(ByVal path As String)
    Dim O() As Byte, h As Long, i As Long, pos As Long, cap As Long
    Dim hd As Long, rawLen As Long, hasRaw As Boolean, nRel As Long

    cap = 20 + 40& * mSecCount + 18& * mSymCount + UBound(StrBuf) + 1 + R1Len + R2Len + RRCount * 10& + 64
    For i = 0 To mSecCount - 1
        If i <> mSecR1 And i <> mSecR2 Then cap = cap + SecSize(i) + 4 + U16(F, 20 + 40& * i + 32) * 10&
    Next
    ReDim O(cap - 1)

    CopyMemory O(0), F(0), 20
    pos = 20 + 40& * mSecCount
    ' raw data
    For i = 0 To mSecCount - 1
        hd = 20 + 40& * i
        CopyMemory O(hd), F(hd), 40
        hasRaw = (SecRaw(i) <> 0) Or i = mSecR1 Or i = mSecR2
        If hasRaw Then
            pos = ((pos + 3) \ 4) * 4
            PutI32 O, hd + 20, pos
            If i = mSecR1 Then
                PutI32 O, hd + 16, R1Len
                CopyMemory O(pos), R1(0), R1Len
                pos = pos + R1Len
            ElseIf i = mSecR2 Then
                PutI32 O, hd + 16, R2Len
                If R2Len > 0 Then CopyMemory O(pos), R2(0), R2Len
                pos = pos + R2Len
            Else
                rawLen = SecSize(i)
                If rawLen > 0 Then CopyMemory O(pos), F(SecRaw(i)), rawLen
                pos = pos + rawLen
            End If
        Else
            PutI32 O, hd + 20, 0
        End If
    Next
    ' relocations
    For i = 0 To mSecCount - 1
        hd = 20 + 40& * i
        If i = mSecR1 Then
            nRel = RRCount
        Else
            nRel = U16(F, hd + 32)
        End If
        If nRel > 0 Then
            PutI32 O, hd + 24, pos
            PutU16 O, hd + 32, nRel
            If i = mSecR1 Then
                CopyMemory O(pos), RR(0), nRel * 10&
            Else
                CopyMemory O(pos), F(I32(F, hd + 24)), nRel * 10&
            End If
            pos = pos + nRel * 10&
        Else
            PutI32 O, hd + 24, 0
            PutU16 O, hd + 32, 0
        End If
    Next
    ' symbols + strings
    PutI32 O, 8, pos
    PutI32 O, 12, mSymCount
    CopyMemory O(pos), SymBuf(0), 18& * mSymCount
    pos = pos + 18& * mSymCount
    CopyMemory O(pos), StrBuf(0), UBound(StrBuf) + 1
    pos = pos + UBound(StrBuf) + 1
    ReDim Preserve O(pos - 1)

    If Len(Dir$(path)) > 0 Then Kill path
    h = FreeFile
    Open path For Binary Access Write As #h
    Put #h, , O
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

Private Sub FindRsrcSections()
    Dim si As Long
    mSecR1 = -1: mSecR2 = -1
    For si = 0 To mSecCount - 1
        Select Case SecName(si)
            Case ".rsrc$01"
                If mSecR1 >= 0 Then Err.Raise vbObjectError + 1, , "more than one .rsrc$01 section"
                mSecR1 = si
            Case ".rsrc$02"
                If mSecR2 >= 0 Then Err.Raise vbObjectError + 1, , "more than one .rsrc$02 section"
                mSecR2 = si
        End Select
    Next
    If mSecR1 < 0 Or mSecR2 < 0 Then Err.Raise vbObjectError + 1, , "no .rsrc$01/.rsrc$02 sections in this object (project without resources?)"
    If SecRaw(mSecR1) = 0 Or SecRaw(mSecR2) = 0 Then Err.Raise vbObjectError + 1, , "empty resource sections"
    If U16(F, 20 + 40& * mSecR2 + 32) <> 0 Then Err.Raise vbObjectError + 1, , "relocations inside .rsrc$02 are not supported"
End Sub

' Index of the static section symbol of .rsrc$02 (value 0); appends one if absent.
Private Function FindSecSym() As Long
    Dim i As Long, p As Long, k As Long, nm As String

    i = 0
    Do While i < mSymCount
        p = i * 18&
        If I16(SymBuf, p + 12) = mSecR2 + 1 And I32(SymBuf, p + 8) = 0 And SymBuf(p + 16) = SYM_STATIC And I32(SymBuf, p) <> 0 Then
            nm = ""
            For k = 0 To 7
                If SymBuf(p + k) = 0 Then Exit For
                nm = nm & Chr$(SymBuf(p + k))
            Next
            If nm = ".rsrc$02" Then FindSecSym = i: Exit Function
        End If
        i = i + 1 + SymBuf(p + 17)
    Loop

    p = mSymCount * 18&                                 ' none found: add "$VRSRC02"
    ReDim Preserve SymBuf(p + 17)
    For k = 0 To 17: SymBuf(p + k) = 0: Next
    nm = "$VRSRC02"
    For k = 1 To 8: SymBuf(p + k - 1) = Asc(Mid$(nm, k, 1)): Next
    PutU16 SymBuf, p + 12, mSecR2 + 1
    SymBuf(p + 16) = SYM_STATIC
    FindSecSym = mSymCount
    mSymCount = mSymCount + 1
End Function

'------------------------------------------------------------------------------
' Resource directory -> flat list of leaves
'------------------------------------------------------------------------------
Private Function ReadUStr(ByVal p As Long) As String
    Dim n As Long, s As String
    n = U16(F, p)
    s = Space$(n)
    If n > 0 Then CopyMemory ByVal StrPtr(s), F(p + 2), n * 2
    ReadUStr = s
End Function

Private Sub DirInfo(ByVal off As Long, ByRef nEnt As Long)
    Dim p As Long
    p = SecRaw(mSecR1) + off
    If off < 0 Or off + 16 > SecSize(mSecR1) Then Err.Raise vbObjectError + 1, , "damaged resource directory"
    nEnt = U16(F, p + 12) + U16(F, p + 14)
    If off + 16 + 8& * nEnt > SecSize(mSecR1) Then Err.Raise vbObjectError + 1, , "damaged resource directory"
End Sub

' Reads entry i of the directory at `off`: its name/id field (string or id) and offset field.
Private Sub DirEntry(ByVal off As Long, ByVal i As Long, ByRef nm As String, ByRef id As Long, ByRef target As Long, ByRef isDir As Boolean)
    Dim p As Long, v As Long
    p = SecRaw(mSecR1) + off + 16 + 8& * i
    v = I32(F, p)
    If (v And &H80000000) <> 0 Then
        v = v And &H7FFFFFFF
        If v + 2 > SecSize(mSecR1) Then Err.Raise vbObjectError + 1, , "damaged resource name"
        nm = ReadUStr(SecRaw(mSecR1) + v)
        id = 0
    Else
        nm = ""
        id = v
    End If
    v = I32(F, p + 4)
    isDir = ((v And &H80000000) <> 0)
    target = v And &H7FFFFFFF
End Sub

Private Sub ParseRsrc()
    Dim nT As Long, nN As Long, nL As Long, t As Long, n As Long, l As Long
    Dim tNm As String, tId As Long, tOff As Long, tDir As Boolean
    Dim nNm As String, nId As Long, nOff As Long, nDir As Boolean
    Dim lNm As String, lId As Long, lOff As Long, lDir As Boolean
    Dim rawR1 As Long, rawR2 As Long, q As Long, sIdx As Long, value As Long, off As Long, sz As Long

    rawR1 = SecRaw(mSecR1): rawR2 = SecRaw(mSecR2)
    CopyMemory mDirHdr(0), F(rawR1), 12
    nRes = 0
    ReDim rTypN(0): ReDim rTypI(0): ReDim rNamN(0): ReDim rNamI(0): ReDim rLang(0)
    ReDim rCP(0): ReDim rRsv(0): ReDim rPool(0): ReDim rSize(0): ReDim rSym(0)
    ReDim rDel(0): ReDim rVB6(0)
    PoolCap = 4096: ReDim Pool(PoolCap - 1): PoolLen = 0

    DirInfo 0, nT
    For t = 0 To nT - 1
        DirEntry 0, t, tNm, tId, tOff, tDir
        If Not tDir Then Err.Raise vbObjectError + 1, , "unsupported resource tree (type entry is not a directory)"
        DirInfo tOff, nN
        For n = 0 To nN - 1
            DirEntry tOff, n, nNm, nId, nOff, nDir
            If Not nDir Then Err.Raise vbObjectError + 1, , "unsupported resource tree (name entry is not a directory)"
            DirInfo nOff, nL
            For l = 0 To nL - 1
                DirEntry nOff, l, lNm, lId, lOff, lDir
                If lDir Or Len(lNm) > 0 Then Err.Raise vbObjectError + 1, , "unsupported resource tree (more than 3 levels or named languages)"
                If lOff + 16 > SecSize(mSecR1) Then Err.Raise vbObjectError + 1, , "damaged resource data entry"
                q = FindReloc(mSecR1, lOff)
                If q < 0 Then Err.Raise vbObjectError + 1, , "resource data entry without relocation"
                If U16(F, q + 8) <> REL_DIR32NB Then Err.Raise vbObjectError + 1, , "unexpected relocation type in .rsrc$01"
                sIdx = I32(F, q + 4)
                If sIdx < 0 Or sIdx >= mSymCount Then Err.Raise vbObjectError + 1, , "bad symbol index in .rsrc$01 relocation"
                If I16(SymBuf, sIdx * 18& + 12) <> mSecR2 + 1 Then Err.Raise vbObjectError + 1, , "resource data is not in .rsrc$02"
                value = I32(SymBuf, sIdx * 18& + 8)
                off = value + I32(F, rawR1 + lOff)
                sz = I32(F, rawR1 + lOff + 4)
                If off < 0 Or sz < 0 Or off + sz > SecSize(mSecR2) Then Err.Raise vbObjectError + 1, , "resource data outside of .rsrc$02"
                AddLeaf tNm, tId, nNm, nId, lId, I32(F, rawR1 + lOff + 8), I32(F, rawR1 + lOff + 12), F, rawR2 + off, sz, sIdx
            Next
        Next
    Next
    If nRes <> U16(F, 20 + 40& * mSecR1 + 32) Then Err.Raise vbObjectError + 1, , "unexpected relocations in .rsrc$01"
End Sub

Private Sub AddLeaf(ByVal tN As String, ByVal tI As Long, ByVal nN As String, ByVal nI As Long, ByVal lang As Long, ByVal cp As Long, ByVal rsv As Long, ByRef src() As Byte, ByVal srcOff As Long, ByVal size As Long, ByVal oldSym As Long)
    ReDim Preserve rTypN(nRes): ReDim Preserve rTypI(nRes): ReDim Preserve rNamN(nRes): ReDim Preserve rNamI(nRes)
    ReDim Preserve rLang(nRes): ReDim Preserve rCP(nRes): ReDim Preserve rRsv(nRes): ReDim Preserve rPool(nRes)
    ReDim Preserve rSize(nRes): ReDim Preserve rSym(nRes): ReDim Preserve rDel(nRes): ReDim Preserve rVB6(nRes)
    rTypN(nRes) = tN: rTypI(nRes) = tI: rNamN(nRes) = nN: rNamI(nRes) = nI
    rLang(nRes) = lang: rCP(nRes) = cp: rRsv(nRes) = rsv: rSize(nRes) = size: rSym(nRes) = oldSym
    rDel(nRes) = False: rVB6(nRes) = False
    If PoolLen + size + 4 > PoolCap Then
        PoolCap = (PoolLen + size + 4) * 2
        ReDim Preserve Pool(PoolCap - 1)
    End If
    rPool(nRes) = PoolLen
    If size > 0 Then CopyMemory Pool(PoolLen), src(srcOff), size
    PoolLen = PoolLen + size
    nRes = nRes + 1
End Sub

'------------------------------------------------------------------------------
' Which resources are the VB6 default icon?
'   RT_GROUP_ICON with the numeric id 1, plus the RT_ICON ids that group lists -
'   unless one of them is also used by another group (e.g. your own "APPICON").
'------------------------------------------------------------------------------
Private Sub FindVB6Icon()
    Dim k As Long, j As Long, c As Long, ids() As Long, nIds As Long, id As Long, p As Long

    vbFound = False: vbGrpLang = 0: vbIcoLang = -1: vbCP = -1
    ReDim ids(0)
    For k = 0 To nRes - 1
        If IsGroupIcon(k) And Len(rNamN(k)) = 0 And rNamI(k) = VB6_GROUP_ID Then
            rVB6(k) = True
            If Not vbFound Then vbGrpLang = rLang(k): vbCP = rCP(k)
            vbFound = True
            If rSize(k) >= 6 Then
                c = U16(Pool, rPool(k) + 4)
                If 6 + 14& * c <= rSize(k) Then
                    For j = 0 To c - 1
                        id = U16(Pool, rPool(k) + 6 + 14& * j + 12)
                        ReDim Preserve ids(nIds): ids(nIds) = id: nIds = nIds + 1
                    Next
                End If
            End If
        End If
    Next
    If Not vbFound Then Exit Sub
    ' ids that another group still needs are not ours
    For k = 0 To nRes - 1
        If IsGroupIcon(k) And Not rVB6(k) And rSize(k) >= 6 Then
            c = U16(Pool, rPool(k) + 4)
            If 6 + 14& * c <= rSize(k) Then
                For j = 0 To c - 1
                    id = U16(Pool, rPool(k) + 6 + 14& * j + 12)
                    For p = 0 To nIds - 1
                        If ids(p) = id Then ids(p) = -1
                    Next
                Next
            End If
        End If
    Next
    For k = 0 To nRes - 1
        If Len(rTypN(k)) = 0 And rTypI(k) = RT_ICON And Len(rNamN(k)) = 0 Then
            For p = 0 To nIds - 1
                If ids(p) = rNamI(k) Then
                    rVB6(k) = True
                    If vbIcoLang < 0 Then vbIcoLang = rLang(k)
                End If
            Next
        End If
    Next
    If vbIcoLang < 0 Then vbIcoLang = vbGrpLang
End Sub

Private Function IsGroupIcon(ByVal k As Long) As Boolean
    IsGroupIcon = (Len(rTypN(k)) = 0 And rTypI(k) = RT_GROUP_ICON)
End Function

Private Function CountGroups() As Long
    Dim k As Long, n As Long
    For k = 0 To nRes - 1
        If rVB6(k) And IsGroupIcon(k) Then n = n + 1
    Next
    CountGroups = n
End Function

Private Function CountLive() As Long
    Dim k As Long, n As Long
    For k = 0 To nRes - 1
        If Not rDel(k) Then n = n + 1
    Next
    CountLive = n
End Function

Private Function TypeText(ByVal k As Long) As String
    If Len(rTypN(k)) > 0 Then TypeText = """" & rTypN(k) & """": Exit Function
    Select Case rTypI(k)
        Case 1: TypeText = "CURSOR"
        Case 2: TypeText = "BITMAP"
        Case 3: TypeText = "ICON"
        Case 4: TypeText = "MENU"
        Case 5: TypeText = "DIALOG"
        Case 6: TypeText = "STRING"
        Case 9: TypeText = "ACCELERATOR"
        Case 10: TypeText = "RCDATA"
        Case 11: TypeText = "MESSAGETABLE"
        Case 12: TypeText = "GROUP_CURSOR"
        Case 14: TypeText = "GROUP_ICON"
        Case 16: TypeText = "VERSION"
        Case 24: TypeText = "MANIFEST"
        Case Else: TypeText = "#" & rTypI(k)
    End Select
End Function

Private Sub ListResources(ByVal path As String)
    Dim k As Long, nm As String
    Out "Resources in " & path & ":"
    For k = 0 To nRes - 1
        If Len(rNamN(k)) > 0 Then nm = """" & rNamN(k) & """" Else nm = "#" & rNamI(k)
        Out "  " & IIf(rVB6(k), "*", " ") & " " & TypeText(k) & " " & nm & "  lang 0x" & Hex$(rLang(k)) & "  " & rSize(k) & " bytes"
    Next
    If vbFound Then
        Out "  (*) = VB6 default icon: replaced / removed by VBIcon"
    Else
        Out "  no VB6 default icon found"
    End If
End Sub

'------------------------------------------------------------------------------
' Rebuilding .rsrc$01 / .rsrc$02
'------------------------------------------------------------------------------
Private Function CmpKey(ByVal sa As String, ByVal ia As Long, ByVal sb As String, ByVal ib As Long) As Long
    Dim na As Boolean, nb As Boolean
    na = (Len(sa) > 0): nb = (Len(sb) > 0)
    If na And Not nb Then CmpKey = -1: Exit Function
    If nb And Not na Then CmpKey = 1: Exit Function
    If na Then
        CmpKey = StrComp(UCase$(sa), UCase$(sb), vbBinaryCompare)
    ElseIf ia < ib Then
        CmpKey = -1
    ElseIf ia > ib Then
        CmpKey = 1
    Else
        CmpKey = 0
    End If
End Function

Private Function CmpLeaf(ByVal a As Long, ByVal b As Long) As Long
    Dim c As Long
    c = CmpKey(rTypN(a), rTypI(a), rTypN(b), rTypI(b))
    If c = 0 Then c = CmpKey(rNamN(a), rNamI(a), rNamN(b), rNamI(b))
    If c = 0 Then
        If rLang(a) < rLang(b) Then
            c = -1
        ElseIf rLang(a) > rLang(b) Then
            c = 1
        End If
    End If
    CmpLeaf = c
End Function

Private Sub PutUStr(ByVal p As Long, ByVal s As String)
    PutU16 R1, p, Len(s)
    If Len(s) > 0 Then CopyMemory R1(p + 2), ByVal StrPtr(s), Len(s) * 2
End Sub

Private Sub PutDirHdr(ByVal p As Long, ByVal nNamed As Long, ByVal nId As Long)
    CopyMemory R1(p), mDirHdr(0), 12
    PutU16 R1, p + 12, nNamed
    PutU16 R1, p + 14, nId
End Sub

Private Sub BuildRsrc()
    Dim ord() As Long, n As Long, i As Long, j As Long, t As Long, g As Long
    Dim tStart() As Long, nStart() As Long, nT As Long, nN As Long, newT As Boolean, newN As Boolean
    Dim tG0() As Long, tNames() As Long, tDir() As Long, nDir() As Long, deOff() As Long
    Dim tStr() As Long, nStr() As Long, d2() As Long, pos As Long, p As Long, k As Long
    Dim secSym As Long, named As Long, ent As Long, first As Long

    ' ---- live leaves, sorted by type / name / language ----
    ReDim ord(nRes)
    For i = 0 To nRes - 1
        If Not rDel(i) Then ord(n) = i: n = n + 1
    Next
    For i = 1 To n - 1
        t = ord(i): j = i - 1
        Do While j >= 0
            If CmpLeaf(ord(j), t) <= 0 Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = t
    Next

    ' ---- group boundaries ----
    ReDim tStart(n + 1): ReDim nStart(n + 1)
    For i = 0 To n - 1
        newT = (i = 0)
        If i > 0 Then newT = (CmpKey(rTypN(ord(i)), rTypI(ord(i)), rTypN(ord(i - 1)), rTypI(ord(i - 1))) <> 0)
        If newT Then tStart(nT) = i: nT = nT + 1
        newN = newT
        If i > 0 And Not newN Then newN = (CmpKey(rNamN(ord(i)), rNamI(ord(i)), rNamN(ord(i - 1)), rNamI(ord(i - 1))) <> 0)
        If newN Then nStart(nN) = i: nN = nN + 1
    Next
    tStart(nT) = n: nStart(nN) = n
    ReDim tG0(nT): ReDim tNames(nT)
    g = 0
    For t = 0 To nT - 1
        tG0(t) = g
        Do While g < nN
            If nStart(g) >= tStart(t + 1) Then Exit Do
            tNames(t) = tNames(t) + 1
            g = g + 1
        Loop
    Next

    ' ---- layout of .rsrc$01: directories, data entries, strings ----
    ReDim tDir(nT): ReDim nDir(nN): ReDim deOff(n): ReDim tStr(nT): ReDim nStr(nN)
    pos = 16 + 8& * nT
    For t = 0 To nT - 1
        tDir(t) = pos: pos = pos + 16 + 8& * tNames(t)
    Next
    For g = 0 To nN - 1
        nDir(g) = pos: pos = pos + 16 + 8& * (nStart(g + 1) - nStart(g))
    Next
    For i = 0 To n - 1
        deOff(i) = pos: pos = pos + 16
    Next
    For t = 0 To nT - 1
        If Len(rTypN(ord(tStart(t)))) > 0 Then tStr(t) = pos: pos = pos + 2 + 2& * Len(rTypN(ord(tStart(t))))
    Next
    For g = 0 To nN - 1
        If Len(rNamN(ord(nStart(g)))) > 0 Then nStr(g) = pos: pos = pos + 2 + 2& * Len(rNamN(ord(nStart(g))))
    Next
    R1Len = ((pos + 3) \ 4) * 4

    ' ---- layout of .rsrc$02 ----
    ReDim d2(n)
    R2Len = 0
    For i = 0 To n - 1
        R2Len = ((R2Len + 3) \ 4) * 4
        d2(i) = R2Len
        R2Len = R2Len + rSize(ord(i))
    Next
    R2Len = ((R2Len + 3) \ 4) * 4

    ReDim R1(R1Len): ReDim R2(R2Len): ReDim RR(10& * n + 10)
    secSym = FindSecSym()

    ' ---- directories ----
    named = 0
    For t = 0 To nT - 1
        If Len(rTypN(ord(tStart(t)))) > 0 Then named = named + 1
    Next
    PutDirHdr 0, named, nT - named
    For t = 0 To nT - 1
        first = ord(tStart(t))
        ent = 16 + 8& * t
        If Len(rTypN(first)) > 0 Then
            PutI32 R1, ent, &H80000000 Or tStr(t)
            PutUStr tStr(t), rTypN(first)
        Else
            PutI32 R1, ent, rTypI(first)
        End If
        PutI32 R1, ent + 4, &H80000000 Or tDir(t)

        named = 0
        For g = tG0(t) To tG0(t) + tNames(t) - 1
            If Len(rNamN(ord(nStart(g)))) > 0 Then named = named + 1
        Next
        PutDirHdr tDir(t), named, tNames(t) - named
        For g = tG0(t) To tG0(t) + tNames(t) - 1
            first = ord(nStart(g))
            ent = tDir(t) + 16 + 8& * (g - tG0(t))
            If Len(rNamN(first)) > 0 Then
                PutI32 R1, ent, &H80000000 Or nStr(g)
                PutUStr nStr(g), rNamN(first)
            Else
                PutI32 R1, ent, rNamI(first)
            End If
            PutI32 R1, ent + 4, &H80000000 Or nDir(g)

            PutDirHdr nDir(g), 0, nStart(g + 1) - nStart(g)
            For i = nStart(g) To nStart(g + 1) - 1
                ent = nDir(g) + 16 + 8& * (i - nStart(g))
                PutI32 R1, ent, rLang(ord(i))
                PutI32 R1, ent + 4, deOff(i)
            Next
        Next
    Next

    ' ---- data entries, data, relocations ----
    For i = 0 To n - 1
        k = ord(i)
        PutI32 R1, deOff(i), d2(i)                     ' addend: offset inside .rsrc$02
        PutI32 R1, deOff(i) + 4, rSize(k)
        PutI32 R1, deOff(i) + 8, rCP(k)
        PutI32 R1, deOff(i) + 12, rRsv(k)
        If rSize(k) > 0 Then CopyMemory R2(d2(i)), Pool(rPool(k)), rSize(k)
        PutI32 RR, 10& * i, deOff(i)
        PutI32 RR, 10& * i + 4, secSym
        PutU16 RR, 10& * i + 8, REL_DIR32NB
    Next
    RRCount = n

    ' ---- symbols: old per-resource symbols are no longer used, park them at 0 ----
    For i = 0 To nRes - 1
        If rSym(i) >= 0 And rSym(i) <> secSym Then PutI32 SymBuf, rSym(i) * 18& + 8, 0
    Next
    ' aux records of the two section symbols carry the section length / reloc count
    i = 0
    Do While i < mSymCount
        p = i * 18&
        If SymBuf(p + 16) = SYM_STATIC And I32(SymBuf, p + 8) = 0 And SymBuf(p + 17) = 1 Then
            If I16(SymBuf, p + 12) = mSecR1 + 1 Then
                PutI32 SymBuf, p + 18, R1Len
                PutU16 SymBuf, p + 22, n
            ElseIf I16(SymBuf, p + 12) = mSecR2 + 1 Then
                PutI32 SymBuf, p + 18, R2Len
                PutU16 SymBuf, p + 22, 0
            End If
        End If
        i = i + 1 + SymBuf(p + 17)
    Loop
End Sub

'------------------------------------------------------------------------------
' Building the new icon resources
'------------------------------------------------------------------------------
Private Function NextIconId(ByVal startId As Long) As Long
    Dim id As Long, k As Long, used As Boolean
    id = startId
    Do
        id = id + 1
        used = False
        For k = 0 To nRes - 1
            If Not rDel(k) And Len(rTypN(k)) = 0 And rTypI(k) = RT_ICON And Len(rNamN(k)) = 0 And rNamI(k) = id Then used = True: Exit For
        Next
    Loop While used
    NextIconId = id
End Function

Private Sub BuildIcons(ByRef sizes() As Long, ByVal nSizes As Long, ByRef depths() As Long, ByVal nDepths As Long)
    Dim grp() As Byte, nImg As Long, i As Long, j As Long, k As Long, id As Long, p As Long
    Dim d() As Byte, sz As Long, S As Long, dp As Long, isPng As Boolean, cp As Long, lang As Long

    nImg = nSizes * nDepths
    ReDim grp(6 + 14& * nImg - 1)
    PutU16 grp, 2, 1                                   ' idType = icon
    PutU16 grp, 4, nImg
    If vbCP >= 0 Then cp = vbCP Else cp = 0
    If vbFound Then lang = vbIcoLang Else lang = 0
    id = VB6_ICON_BASE - 1                             ' NextIconId returns the first free id above this
    For i = 0 To nSizes - 1
        For j = 0 To nDepths - 1
            S = sizes(i): dp = depths(j)
            isPng = False
            d = MakeImage(S, dp, isPng)
            sz = UBound(d) + 1
            id = NextIconId(id)
            AddLeaf "", RT_ICON, "", id, lang, cp, 0, d, 0, sz, -1
            p = 6 + 14& * k
            If S < 256 Then grp(p) = S
            If S < 256 Then grp(p + 1) = S
            If dp = 1 Then grp(p + 2) = 2
            If dp = 4 Then grp(p + 2) = 16
            PutU16 grp, p + 4, 1                       ' planes
            PutU16 grp, p + 6, dp                      ' bit count
            PutI32 grp, p + 8, sz
            PutU16 grp, p + 12, id
            Out "  " & S & "x" & S & "  " & dp & " bit  " & sz & " bytes" & IIf(isPng, "  (PNG)", "") & "  -> RT_ICON #" & id
            k = k + 1
        Next
    Next
    If vbFound Then lang = vbGrpLang Else lang = 0
    AddLeaf "", RT_GROUP_ICON, "", VB6_GROUP_ID, lang, cp, 0, grp, 0, UBound(grp) + 1, -1
End Sub

' Renders the master image at S x S and converts it to the DIB (or PNG) that goes into RT_ICON.
Private Function MakeImage(ByVal S As Long, ByVal depth As Long, ByRef isPng As Boolean) As Byte()
    Dim bmp As Long, nPal As Long
    Dim pr() As Long, pg() As Long, pb() As Long, idx() As Byte

    If S <> mPxS Then                                  ' the depths of one size follow each other: scale only once
        bmp = ScaleTo(S)
        ReadPixels bmp, S, mPx
        GdipDisposeImage bmp
        mPxS = S
    End If

    If depth = 32 And S >= 256 Then
        bmp = BitmapFromPixels(S, mPx)
        MakeImage = EncodePng(bmp)
        GdipDisposeImage bmp
        isPng = True
    ElseIf depth >= 24 Then
        MakeImage = BuildDib(S, depth, mPx, 0, pr, pg, pb, idx)
    Else
        If depth = 1 Then nPal = 2
        If depth = 4 Then nPal = 16
        If depth = 8 Then nPal = 256
        BuildPalette depth, S, mPx, nPal, pr, pg, pb
        ReDim idx(S * S - 1)
        Atkinson mPx, S, nPal, pr, pg, pb, idx
        MakeImage = BuildDib(S, depth, mPx, nPal, pr, pg, pb, idx)
    End If
End Function

' S x S top-down 32 bit BGRA, straight (not premultiplied) alpha
Private Sub ReadPixels(ByVal bmp As Long, ByVal S As Long, ByRef px() As Byte)
    Dim rc As GpRect, bd As GpBitmapData, y As Long
    rc.X = 0: rc.Y = 0: rc.W = S: rc.H = S
    ReDim px(S * S * 4 - 1)
    If GdipBitmapLockBits(bmp, rc, 1&, PF_32ARGB, bd) <> 0 Then Err.Raise vbObjectError + 1, , "GDI+ LockBits failed"
    For y = 0 To S - 1
        CopyMemory px(y * S * 4&), ByVal bd.Scan0 + y * bd.Stride, S * 4&
    Next
    GdipBitmapUnlockBits bmp, bd
End Sub

' GDI+ bitmap that reads straight from the px() array (keep px alive while it is used)
Private Function BitmapFromPixels(ByVal S As Long, ByRef px() As Byte) As Long
    Dim bmp As Long
    If GdipCreateBitmapFromScan0(S, S, S * 4&, PF_32ARGB, VarPtr(px(0)), bmp) <> 0 Then Err.Raise vbObjectError + 1, , "GDI+ CreateBitmapFromScan0 failed"
    BitmapFromPixels = bmp
End Function

' Fit (or stretch) the master into an S x S premultiplied bitmap
Private Function ScaleTo(ByVal S As Long) As Long
    Dim bmp As Long, g As Long, ia As Long, dx As Long, dy As Long, dw As Long, dh As Long
    Dim sc As Double, sx As Double, sy As Double

    If GdipCreateBitmapFromScan0(S, S, 0, PF_32PARGB, 0, bmp) <> 0 Then Err.Raise vbObjectError + 1, , "GDI+ CreateBitmapFromScan0 failed"
    GdipGetImageGraphicsContext bmp, g
    GdipSetCompositingMode g, 1                        ' SourceCopy
    GdipSetInterpolationMode g, 7                      ' HighQualityBicubic
    GdipSetPixelOffsetMode g, 4                        ' Half
    GdipGraphicsClear g, 0
    sx = S / mSrcW: sy = S / mSrcH
    If mStretch Then
        dw = S: dh = S
    Else
        If sx < sy Then sc = sx Else sc = sy
        dw = CLng(mSrcW * sc): dh = CLng(mSrcH * sc)
        If dw < 1 Then dw = 1
        If dh < 1 Then dh = 1
        If dw > S Then dw = S
        If dh > S Then dh = S
    End If
    dx = (S - dw) \ 2: dy = (S - dh) \ 2
    GdipCreateImageAttributes ia
    GdipSetImageAttributesWrapMode ia, 3, 0, 0         ' TileFlipXY: no faded edges
    If GdipDrawImageRectRectI(g, mSrc, dx, dy, dw, dh, 0, 0, mSrcW, mSrcH, 2, ia, 0, 0) <> 0 Then Err.Raise vbObjectError + 1, , "GDI+ DrawImage failed"
    GdipDisposeImageAttributes ia
    GdipDeleteGraphics g
    ScaleTo = bmp
End Function

Private Function EncodePng(ByVal bmp As Long) As Byte()
    Dim stm As Long, hMem As Long, n As Long, p As Long, tmp() As Byte, pos As Long, ln As Long
    Dim clsid As GpGuid, found As Boolean

    clsid.Data1 = &H557CF406: clsid.Data2 = &H1A04: clsid.Data3 = &H11D3     ' image/png encoder
    clsid.Data4(0) = &H9A: clsid.Data4(1) = &H73: clsid.Data4(2) = 0: clsid.Data4(3) = 0
    clsid.Data4(4) = &HF8: clsid.Data4(5) = &H1E: clsid.Data4(6) = &HF3: clsid.Data4(7) = &H2E
    If CreateStreamOnHGlobal(0&, 1&, stm) <> 0 Then Err.Raise vbObjectError + 1, , "CreateStreamOnHGlobal failed"
    If GdipSaveImageToStream(bmp, stm, clsid, 0&) <> 0 Then ReleaseStream stm: Err.Raise vbObjectError + 1, , "GDI+ PNG encoder failed"
    GetHGlobalFromStream stm, hMem
    n = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim tmp(n - 1)
    CopyMemory tmp(0), ByVal p, n
    GlobalUnlock hMem
    ReleaseStream stm

    ' the HGLOBAL can be bigger than the stream: cut after the IEND chunk
    pos = 8
    Do While pos + 12 <= n
        ln = CLng(tmp(pos)) * 16777216 + CLng(tmp(pos + 1)) * 65536 + CLng(tmp(pos + 2)) * 256& + tmp(pos + 3)
        If tmp(pos + 4) = 73 And tmp(pos + 5) = 69 And tmp(pos + 6) = 78 And tmp(pos + 7) = 68 Then found = True: Exit Do
        pos = pos + 12 + ln
    Loop
    If Not found Then Err.Raise vbObjectError + 1, , "PNG encoder returned damaged data"
    ReDim Preserve tmp(pos + 12 - 1)
    EncodePng = tmp
End Function

Private Sub ReleaseStream(ByVal stm As Long)
    Dim unk As IUnknown
    If stm = 0 Then Exit Sub
    CopyMemory unk, stm, 4                             ' take the pointer without AddRef ...
    Set unk = Nothing                                  ' ... and let VB do the final Release
End Sub

'------------------------------------------------------------------------------
' Palette + Atkinson dithering
'------------------------------------------------------------------------------
Private Sub BuildPalette(ByVal depth As Long, ByVal S As Long, ByRef px() As Byte, ByVal nPal As Long, ByRef pr() As Long, ByRef pg() As Long, ByRef pb() As Long)
    Dim i As Long, ok As Boolean, bmp As Long, buf() As Byte, cnt As Long
    Dim hDC As Long, hPal As Long, pe() As Byte, got As Long
    Dim best As Long, bd As Long, d As Long

    ReDim pr(nPal - 1): ReDim pg(nPal - 1): ReDim pb(nPal - 1)
    If depth = 1 Then
        pr(1) = 255: pg(1) = 255: pb(1) = 255
        Exit Sub
    End If

    ' 1) GDI+ optimal (median cut) palette computed from the image itself
    bmp = BitmapFromPixels(S, px)
    ReDim buf(8 + 4& * nPal - 1)
    PutI32 buf, 4, nPal
    If GdipInitializePalette(buf(0), 1, nPal, 0, bmp) = 0 Then
        cnt = I32(buf, 4)
        If cnt >= 2 And cnt <= nPal Then
            For i = 0 To cnt - 1
                pb(i) = buf(8 + 4& * i): pg(i) = buf(9 + 4& * i): pr(i) = buf(10 + 4& * i)
            Next
            ok = True
        End If
    End If
    GdipDisposeImage bmp

    If Not ok And depth = 8 Then
        ' 2) GDI halftone palette
        hDC = GetDC(0)
        hPal = CreateHalftonePalette(hDC)
        If hPal <> 0 Then
            ReDim pe(1023)
            got = GetPaletteEntries(hPal, 0, 256, pe(0))
            If got >= 2 Then
                For i = 0 To got - 1
                    pr(i) = pe(4& * i): pg(i) = pe(4& * i + 1): pb(i) = pe(4& * i + 2)
                Next
                ok = True
            End If
            DeleteObject hPal
        End If
        ReleaseDC 0, hDC
    End If
    If Not ok Then
        ' 3) fixed palettes: VGA 16 colours / 3-3-2 cube
        For i = 0 To nPal - 1
            If nPal = 16 Then
                If i = 7 Then
                    pr(i) = 192: pg(i) = 192: pb(i) = 192
                ElseIf i = 8 Then
                    pr(i) = 128: pg(i) = 128: pb(i) = 128
                Else
                    If i < 8 Then d = 128 Else d = 255
                    If i And 1 Then pr(i) = d
                    If i And 2 Then pg(i) = d
                    If i And 4 Then pb(i) = d
                End If
            Else
                pr(i) = ((i \ 32) And 7) * 255 \ 7
                pg(i) = ((i \ 4) And 7) * 255 \ 7
                pb(i) = (i And 3) * 255 \ 3
            End If
        Next
    End If

    ' transparent pixels are XORed onto the screen, so the palette needs a real black
    best = 0: bd = &H7FFFFFFF
    For i = 0 To nPal - 1
        d = pr(i) * pr(i) + pg(i) * pg(i) + pb(i) * pb(i)
        If d < bd Then bd = d: best = i
    Next
    If bd > 0 Then pr(best) = 0: pg(best) = 0: pb(best) = 0
End Sub

' Atkinson: 1/8 of the quantisation error goes to the 6 neighbours
'            .  X  1  1
'            1  1  1  .
'            .  1  .  .       (only 6/8 of the error is spread - highlights and shadows stay clean)
Private Sub Atkinson(ByRef px() As Byte, ByVal S As Long, ByVal nPal As Long, ByRef pr() As Long, ByRef pg() As Long, ByRef pb() As Long, ByRef idx() As Byte)
    Dim eR() As Single, eG() As Single, eB() As Single
    Dim W As Long, x As Long, y As Long, p As Long, e As Long, i As Long, best As Long, blk As Long
    Dim r As Single, g As Single, b As Single, rl As Long, gl As Long, bl As Long
    Dim dr As Long, dg As Long, db As Long, d As Long, bd As Long
    Dim sr As Single, sg As Single, sb As Single

    W = S + 4
    ReDim eR(W * (S + 2)): ReDim eG(W * (S + 2)): ReDim eB(W * (S + 2))

    bd = &H7FFFFFFF                                    ' index of the palette's black
    For i = 0 To nPal - 1
        d = pr(i) * pr(i) + pg(i) * pg(i) + pb(i) * pb(i)
        If d < bd Then bd = d: blk = i
    Next

    For y = 0 To S - 1
        For x = 0 To S - 1
            p = (y * S + x) * 4
            If px(p + 3) < 128 Then
                idx(y * S + x) = blk                   ' transparent: AND mask is set, XOR part black
            Else
                e = y * W + x + 2
                r = px(p + 2) + eR(e): g = px(p + 1) + eG(e): b = px(p) + eB(e)
                If r < 0 Then r = 0
                If r > 255 Then r = 255
                If g < 0 Then g = 0
                If g > 255 Then g = 255
                If b < 0 Then b = 0
                If b > 255 Then b = 255
                rl = CLng(r): gl = CLng(g): bl = CLng(b)
                bd = &H7FFFFFFF
                For i = 0 To nPal - 1
                    dr = rl - pr(i): dg = gl - pg(i): db = bl - pb(i)
                    d = dr * dr + dg * dg + db * db
                    If d < bd Then
                        bd = d: best = i
                        If d = 0 Then Exit For
                    End If
                Next
                idx(y * S + x) = best
                sr = (r - pr(best)) / 8: sg = (g - pg(best)) / 8: sb = (b - pb(best)) / 8
                eR(e + 1) = eR(e + 1) + sr: eG(e + 1) = eG(e + 1) + sg: eB(e + 1) = eB(e + 1) + sb
                eR(e + 2) = eR(e + 2) + sr: eG(e + 2) = eG(e + 2) + sg: eB(e + 2) = eB(e + 2) + sb
                eR(e + W - 1) = eR(e + W - 1) + sr: eG(e + W - 1) = eG(e + W - 1) + sg: eB(e + W - 1) = eB(e + W - 1) + sb
                eR(e + W) = eR(e + W) + sr: eG(e + W) = eG(e + W) + sg: eB(e + W) = eB(e + W) + sb
                eR(e + W + 1) = eR(e + W + 1) + sr: eG(e + W + 1) = eG(e + W + 1) + sg: eB(e + W + 1) = eB(e + W + 1) + sb
                eR(e + 2 * W) = eR(e + 2 * W) + sr: eG(e + 2 * W) = eG(e + 2 * W) + sg: eB(e + 2 * W) = eB(e + 2 * W) + sb
            End If
        Next
    Next
End Sub

' ICO image data: BITMAPINFOHEADER (height = 2*S), palette, XOR bitmap, AND mask - rows bottom up.
Private Function BuildDib(ByVal S As Long, ByVal depth As Long, ByRef px() As Byte, ByVal nPal As Long, ByRef pr() As Long, ByRef pg() As Long, ByRef pb() As Long, ByRef idx() As Byte) As Byte()
    Dim o() As Byte, xs As Long, ms As Long, pal As Long, tot As Long, i As Long
    Dim xBase As Long, mBase As Long, x As Long, y As Long, sy As Long, p As Long, q As Long, m As Long, clear As Boolean

    xs = ((S * depth + 31) \ 32) * 4
    ms = ((S + 31) \ 32) * 4
    If depth <= 8 Then pal = nPal * 4
    tot = 40 + pal + (xs + ms) * S
    ReDim o(tot - 1)
    PutI32 o, 0, 40
    PutI32 o, 4, S
    PutI32 o, 8, S * 2
    PutU16 o, 12, 1
    PutU16 o, 14, depth
    PutI32 o, 20, (xs + ms) * S
    If depth <= 8 Then
        For i = 0 To nPal - 1
            o(40 + 4& * i) = pb(i): o(41 + 4& * i) = pg(i): o(42 + 4& * i) = pr(i)
        Next
    End If
    xBase = 40 + pal
    mBase = xBase + xs * S
    For y = 0 To S - 1
        sy = S - 1 - y
        For x = 0 To S - 1
            p = (y * S + x) * 4
            clear = (px(p + 3) < 128)
            If clear Then
                m = mBase + sy * ms + x \ 8
                o(m) = o(m) Or BitMask(x And 7)
            End If
            q = xBase + sy * xs
            Select Case depth
                Case 32
                    o(q + x * 4) = px(p): o(q + x * 4 + 1) = px(p + 1): o(q + x * 4 + 2) = px(p + 2): o(q + x * 4 + 3) = px(p + 3)
                Case 24
                    If Not clear Then o(q + x * 3) = px(p): o(q + x * 3 + 1) = px(p + 1): o(q + x * 3 + 2) = px(p + 2)
                Case 8
                    o(q + x) = idx(y * S + x)
                Case 4
                    If (x And 1) = 0 Then
                        o(q + x \ 2) = idx(y * S + x) * 16
                    Else
                        o(q + x \ 2) = o(q + x \ 2) Or idx(y * S + x)
                    End If
                Case 1
                    If idx(y * S + x) <> 0 Then o(q + x \ 8) = o(q + x \ 8) Or BitMask(x And 7)
            End Select
        Next
    Next
    BuildDib = o
End Function

'------------------------------------------------------------------------------
' Source image
'------------------------------------------------------------------------------
Private Sub StartGdip()
    Dim si As GdiplusStartupInput
    si.GdiplusVersion = 1
    If GdiplusStartup(mToken, si, 0&) <> 0 Then mToken = 0: Err.Raise vbObjectError + 1, , "GDI+ could not be started"
End Sub

Private Sub LoadSource(ByVal path As String)
    Dim b() As Byte, h As Long, n As Long

    If Len(Dir$(path)) = 0 Then Err.Raise vbObjectError + 1, , "file not found: " & path
    n = FileLen(path)
    If n < 8 Then Err.Raise vbObjectError + 1, , "not an image file: " & path
    ReDim b(n - 1)
    h = FreeFile
    Open path For Binary Access Read As #h
    Get #h, , b
    Close #h

    If b(0) = 0 And b(1) = 0 And b(2) = 1 And b(3) = 0 Then
        LoadIco b, n
    Else
        LoadBytes b, 0, n
    End If
    If mSrc = 0 Then Err.Raise vbObjectError + 1, , "cannot decode image (need JPEG, PNG, BMP or ICO): " & path
End Sub

' ICO: take the biggest / deepest frame
Private Sub LoadIco(ByRef b() As Byte, ByVal n As Long)
    Dim cnt As Long, i As Long, e As Long, w As Long, h As Long, score As Long, best As Long, bestScore As Long
    Dim off As Long, sz As Long, hIcon As Long, img As Long, bw As Long, bh As Long, bOff As Long, bSz As Long

    cnt = U16(b, 4)
    best = -1: bestScore = -1
    For i = 0 To cnt - 1
        e = 6 + 16& * i
        If e + 16 > n Then Exit For
        w = b(e): If w = 0 Then w = 256
        h = b(e + 1): If h = 0 Then h = 256
        sz = I32(b, e + 8): off = I32(b, e + 12)
        If off >= 0 And sz > 0 And off + sz <= n Then
            score = w * h * 64& + U16(b, e + 6)
            If score > bestScore Then bestScore = score: best = i: bw = w: bh = h: bOff = off: bSz = sz
        End If
    Next
    If best < 0 Then Err.Raise vbObjectError + 1, , "ICO file has no usable image"

    If b(bOff) = &H89 And b(bOff + 1) = &H50 And b(bOff + 2) = &H4E And b(bOff + 3) = &H47 Then
        LoadBytes b, bOff, bSz                         ' PNG compressed frame -> GDI+
    Else
        hIcon = CreateIconFromResourceEx(b(bOff), bSz, 1&, &H30000, bw, bh, 0&)     ' DIB frame -> GDI
        If hIcon = 0 Then Err.Raise vbObjectError + 1, , "cannot decode the ICO frame"
        If GdipCreateBitmapFromHICON(hIcon, img) <> 0 Then DestroyIcon hIcon: Err.Raise vbObjectError + 1, , "GDI+ cannot read the icon"
        DestroyIcon hIcon
        MakeMaster img
        GdipDisposeImage img
    End If
End Sub

' JPEG / PNG / BMP ... straight from memory through an IStream
Private Sub LoadBytes(ByRef b() As Byte, ByVal off As Long, ByVal n As Long)
    Dim hMem As Long, p As Long, stm As Long, img As Long, st As Long

    hMem = GlobalAlloc(2&, n)                          ' GMEM_MOVEABLE
    If hMem = 0 Then Err.Raise vbObjectError + 1, , "out of memory"
    p = GlobalLock(hMem)
    CopyMemory ByVal p, b(off), n
    GlobalUnlock hMem
    If CreateStreamOnHGlobal(hMem, 1&, stm) <> 0 Then GlobalFree hMem: Err.Raise vbObjectError + 1, , "CreateStreamOnHGlobal failed"
    st = GdipCreateBitmapFromStream(stm, img)
    If st = 0 Then
        FixOrientation img
        MakeMaster img
        GdipDisposeImage img
    End If
    ReleaseStream stm
End Sub

' JPEG EXIF orientation (tag 0x0112): rotate so that "up" is up
Private Sub FixOrientation(ByVal img As Long)
    Dim sz As Long, buf() As Byte, ptr As Long, o As Integer, rf As Long

    If GdipGetPropertyItemSize(img, &H112, sz) <> 0 Then Exit Sub
    If sz < 18 Then Exit Sub
    ReDim buf(sz - 1)
    If GdipGetPropertyItem(img, &H112, sz, buf(0)) <> 0 Then Exit Sub
    ptr = I32(buf, 8 + 4)                              ' PropertyItem.value
    If ptr = 0 Then Exit Sub
    CopyMemory o, ByVal ptr, 2
    Select Case o
        Case 2: rf = 4                                 ' mirrored
        Case 3: rf = 2                                 ' 180
        Case 4: rf = 6
        Case 5: rf = 5                                 ' transpose = Rotate90FlipX
        Case 6: rf = 1                                 ' 90 cw
        Case 7: rf = 7                                 ' transverse = Rotate270FlipX
        Case 8: rf = 3                                 ' 270 cw
        Case Else: Exit Sub
    End Select
    GdipImageRotateFlip img, rf
End Sub

' Copy any decoded image into a premultiplied ARGB bitmap: resampling is then alpha correct
Private Sub MakeMaster(ByVal img As Long)
    Dim w As Long, h As Long, g As Long

    GdipGetImageWidth img, w
    GdipGetImageHeight img, h
    If w < 1 Or h < 1 Or w > 32767 Or h > 32767 Then Err.Raise vbObjectError + 1, , "bad image size"
    If GdipCreateBitmapFromScan0(w, h, 0, PF_32PARGB, 0, mSrc) <> 0 Then mSrc = 0: Err.Raise vbObjectError + 1, , "GDI+ CreateBitmapFromScan0 failed"
    GdipGetImageGraphicsContext mSrc, g
    GdipSetCompositingMode g, 1
    GdipSetInterpolationMode g, 5                      ' NearestNeighbor: exact 1:1 copy
    GdipSetPixelOffsetMode g, 4
    If GdipDrawImageRectRectI(g, img, 0, 0, w, h, 0, 0, w, h, 2, 0, 0, 0) <> 0 Then
        GdipDeleteGraphics g
        GdipDisposeImage mSrc: mSrc = 0
        Err.Raise vbObjectError + 1, , "GDI+ cannot draw the source image"
    End If
    GdipDeleteGraphics g
    mSrcW = w: mSrcH = h
End Sub

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
        WriteFile mStdOut, b(0), UBound(b) + 1, w, 0&          ' console, file or pipe
    End If
End Sub

Private Sub Finish(ByVal rc As Long)
    If mSrc <> 0 Then GdipDisposeImage mSrc: mSrc = 0
    If mToken <> 0 Then GdiplusShutdown mToken: mToken = 0
    If Not InIDE() Then ExitProcess rc
End Sub

Private Function InIDE() As Boolean
    Debug.Assert MarkIDE(InIDE)                        ' only executed inside the IDE
End Function

Private Function MarkIDE(ByRef b As Boolean) As Boolean
    b = True
    MarkIDE = True
End Function

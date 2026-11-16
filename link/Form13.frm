VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Linker"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_cfg 
      Caption         =   "Config:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   735
   End
   Begin VB.ComboBox cbo_cfg 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2900
   End
   Begin VB.CommandButton btn_cfgsave 
      Caption         =   "Save"
      Height          =   345
      Left            =   3900
      TabIndex        =   2
      Top             =   105
      Width           =   900
   End
   Begin VB.CommandButton btn_cfgsaveas 
      Caption         =   "Save as..."
      Height          =   345
      Left            =   4860
      TabIndex        =   3
      Top             =   105
      Width           =   1050
   End
   Begin VB.CommandButton btn_cfgdel 
      Caption         =   "Delete"
      Height          =   345
      Left            =   5970
      TabIndex        =   4
      Top             =   105
      Width           =   900
   End
   Begin VB.CommandButton btn_tools 
      Caption         =   "Tools..."
      Height          =   345
      Left            =   6930
      TabIndex        =   5
      Top             =   105
      Width           =   960
   End
   Begin VB.OptionButton opt_tab 
      Caption         =   "General"
      Height          =   345
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.OptionButton opt_tab 
      Caption         =   "DLL"
      Height          =   345
      Index           =   1
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   600
      Width           =   1500
   End
   Begin VB.Frame fra_tab 
      Caption         =   ""
      Height          =   4680
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   7800
      Begin VB.CheckBox check_static 
         Caption         =   "Embed MSVBVM60.DLL into build (VBStatic)"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   0   'Unchecked
         Width           =   7400
      End
      Begin VB.Label lbl_static_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   465
         Width           =   800
      End
      Begin VB.TextBox txt_static_obj 
         Height          =   315
         Left            =   1300
         TabIndex        =   2
         Top             =   420
         Width           =   4300
      End
      Begin VB.CommandButton btn_static_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5700
         TabIndex        =   3
         Top             =   420
         Width           =   600
      End
      Begin VB.CommandButton btn_static_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   6350
         TabIndex        =   4
         Top             =   420
         Width           =   450
      End
      Begin VB.Label lbl_static_state 
         Caption         =   ""
         Height          =   220
         Left            =   1300
         TabIndex        =   5
         Top             =   760
         Width           =   5500
      End
      Begin VB.CheckBox check_autofix 
         Caption         =   "Adjust link args for VBStatic (entry point, drop msvbvm60.lib, /NODEFAULTLIB)"
         Height          =   240
         Left            =   480
         TabIndex        =   6
         Top             =   1000
         Value           =   1   'Checked
         Width           =   7200
      End
      Begin VB.CheckBox check_extern 
         Caption         =   "Support for external symbols (VBExtern)"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Value           =   0   'Unchecked
         Width           =   7400
      End
      Begin VB.Label lbl_extern_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1725
         Width           =   800
      End
      Begin VB.TextBox txt_extern_obj 
         Height          =   315
         Left            =   1300
         TabIndex        =   9
         Top             =   1680
         Width           =   4300
      End
      Begin VB.CommandButton btn_extern_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5700
         TabIndex        =   10
         Top             =   1680
         Width           =   600
      End
      Begin VB.CommandButton btn_extern_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   6350
         TabIndex        =   11
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label lbl_extern_state 
         Caption         =   ""
         Height          =   220
         Left            =   1300
         TabIndex        =   12
         Top             =   2020
         Width           =   5500
      End
      Begin VB.CheckBox check_externall 
         Caption         =   "Patch every .obj that contains Declare thunks (not only main.obj)"
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   2260
         Value           =   0   'Unchecked
         Width           =   7200
      End
      Begin VB.Label lbl_tool 
         Caption         =   "Arguments for:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2705
         Width           =   1300
      End
      Begin VB.ComboBox cbo_tool 
         Height          =   315
         Left            =   1450
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2660
         Width           =   2400
      End
      Begin VB.ListBox args_list 
         Height          =   1500
         Index           =   0
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   3040
         Visible         =   -1  'True
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1500
         Index           =   1
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   3040
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1500
         Index           =   2
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   3040
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.CommandButton btn_add 
         Caption         =   "+"
         Height          =   285
         Left            =   7000
         TabIndex        =   19
         Top             =   3040
         Width           =   720
      End
      Begin VB.CommandButton btn_remove 
         Caption         =   "-"
         Height          =   285
         Left            =   7000
         TabIndex        =   20
         Top             =   3340
         Width           =   720
      End
      Begin VB.CommandButton btn_addfile 
         Caption         =   "File..."
         Height          =   285
         Left            =   7000
         TabIndex        =   21
         Top             =   3640
         Width           =   720
      End
      Begin VB.CommandButton btn_up 
         Caption         =   "Up"
         Height          =   285
         Left            =   7000
         TabIndex        =   22
         Top             =   3940
         Width           =   720
      End
      Begin VB.CommandButton btn_down 
         Caption         =   "Down"
         Height          =   285
         Left            =   7000
         TabIndex        =   23
         Top             =   4240
         Width           =   720
      End
   End
   Begin VB.Frame fra_tab 
      Caption         =   ""
      Height          =   4680
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      Begin VB.OptionButton opt_dll 
         Caption         =   "ActiveX DLL"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   150
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.OptionButton opt_dll 
         Caption         =   "Standard DLL"
         Height          =   255
         Index           =   1
         Left            =   2300
         TabIndex        =   1
         Top             =   150
         Width           =   2000
      End
      Begin VB.Label lbl_funcs 
         Caption         =   "Functions (filter):"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   610
         Width           =   1500
      End
      Begin VB.TextBox txt_filter 
         Height          =   315
         Left            =   1700
         TabIndex        =   3
         Top             =   570
         Width           =   1720
      End
      Begin VB.Label lbl_exports 
         Caption         =   "Exports (.DEF)"
         Height          =   255
         Left            =   4200
         TabIndex        =   4
         Top             =   610
         Width           =   3000
      End
      Begin VB.ListBox lst_funcs 
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   3300
      End
      Begin VB.ListBox lst_exports 
         Height          =   2400
         Left            =   4200
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   960
         Width           =   3500
      End
      Begin VB.CommandButton btn_exp_add 
         Caption         =   ">"
         Height          =   345
         Left            =   3500
         TabIndex        =   7
         Top             =   1100
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_addall 
         Caption         =   ">>"
         Height          =   345
         Left            =   3500
         TabIndex        =   8
         Top             =   1550
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_remove 
         Caption         =   "<"
         Height          =   345
         Left            =   3500
         TabIndex        =   9
         Top             =   2000
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_clear 
         Caption         =   "<<"
         Height          =   345
         Left            =   3500
         TabIndex        =   10
         Top             =   2450
         Width           =   600
      End
      Begin VB.CommandButton btn_rescan 
         Caption         =   "Rescan OBJs"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   3460
         Width           =   1300
      End
      Begin VB.CommandButton btn_exp_manual 
         Caption         =   "Add name..."
         Height          =   345
         Left            =   4200
         TabIndex        =   12
         Top             =   3460
         Width           =   1300
      End
      Begin VB.Label lbl_def 
         Caption         =   ".DEF file:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3905
         Width           =   900
      End
      Begin VB.TextBox txt_def 
         Height          =   315
         Left            =   1050
         TabIndex        =   14
         Top             =   3860
         Width           =   4700
      End
      Begin VB.CommandButton btn_def_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   5800
         TabIndex        =   15
         Top             =   3860
         Width           =   375
      End
      Begin VB.CommandButton btn_createdef 
         Caption         =   "Create .DEF"
         Height          =   345
         Left            =   6300
         TabIndex        =   16
         Top             =   3845
         Width           =   1400
      End
      Begin VB.CheckBox check_autodef 
         Caption         =   "Regenerate the .DEF automatically on compile (Standard DLL)"
         Height          =   240
         Left            =   1050
         TabIndex        =   17
         Top             =   4260
         Value           =   1   'Checked
         Width           =   6000
      End
   End
   Begin VB.CheckBox check_cmd 
      Caption         =   "Run in CMD"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5730
      Value           =   0   'Unchecked
      Width           =   1500
   End
   Begin VB.CheckBox check_clean 
      Caption         =   "Delete patched OBJs after a successful link"
      Height          =   255
      Left            =   1700
      TabIndex        =   11
      Top             =   5730
      Value           =   1   'Checked
      Width           =   4000
   End
   Begin VB.Label lbl_status 
      Caption         =   ""
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   6030
      Width           =   7800
   End
   Begin VB.CommandButton btn_preview 
      Caption         =   "Copy commands"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6360
      Width           =   1500
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "Compile"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btn_exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   15
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Menu mnu_tools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnu_tool_link 
         Caption         =   "LINK2.EXE"
      End
      Begin VB.Menu mnu_tool_static 
         Caption         =   "VBStatic.exe"
      End
      Begin VB.Menu mnu_tool_extern 
         Caption         =   "VBExtern.exe"
      End
      Begin VB.Menu mnu_tool_runtime 
         Caption         =   "MSVBVM60.DLL"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' Linker front end (LINK.EXE replacement for VB6) with VBStatic + VBExtern
'
' Build pipeline (every patched OBJ is written to a NEW file via /OUT:, the
' original OBJ is never modified, so Compile can be pressed again and again):
'
'   1. VBExtern  <obj> /OUT:<obj>.vbx.obj      (Declare ... Lib "<extern>")
'   2. VBStatic  [/DLL:x] /OUT:<obj>.vbs.obj <msvbvm60.dll> <obj> [chain]
'   3. LINK2.EXE with the patched OBJ(s) substituted into the argument list
'
' Build configs are stored in "vblinker.ini" in the directory of main.obj,
' tool paths in "vblinker.ini" next to this exe.
'==============================================================================

'---- Win32 ---------------------------------------------------------------------
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const STARTF_USESHOWWINDOW As Long = 1
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMINNOACTIVE As Long = 7
Private Const WAIT_TIMEOUT As Long = &H102&

Private Const OFN_HIDEREADONLY As Long = &H4&
Private Const OFN_NOCHANGEDIR As Long = &H8&
Private Const OFN_OVERWRITEPROMPT As Long = &H2&
Private Const OFN_PATHMUSTEXIST As Long = &H800&
Private Const OFN_FILEMUSTEXIST As Long = &H1000&
Private Const OFN_ALLOWMULTISELECT As Long = &H200&
Private Const OFN_EXPLORER As Long = &H80000

'---- constants -------------------------------------------------------------------
Private Const TOOL_LINK As Long = 0
Private Const TOOL_STATIC As Long = 1
Private Const TOOL_EXTERN As Long = 2

Private Const FLAG_USER As Long = 0        'ListBox.ItemData: added / edited by the user
Private Const FLAG_CMD As Long = 1         'ListBox.ItemData: came from the command line

'---- state -----------------------------------------------------------------------
Private mLoading As Boolean
Private mDirty As Boolean
Private mExitCode As Long

Private mCmdArgs() As String               'arguments passed by VB6
Private mCmdCount As Long
Private mHidden() As String                'command line args the user removed / edited away
Private mHiddenCount As Long

Private mCfgDir As String
Private mCfgFile As String
Private mCurCfg As String

Private mToolLink As String
Private mToolStatic As String
Private mToolExtern As String
Private mRuntime As String

Private mFnName() As String                'functions found in the OBJs
Private mFnScope() As String
Private mFnDecor() As String
Private mFnCount As Long

Private mMapOrig() As String               'original OBJ -> patched OBJ
Private mMapNew() As String
Private mMapCount As Long
Private mTemp() As String                  'patched OBJs to delete after the link
Private mTempCount As Long

Private mOb() As Byte                      'COFF object currently loaded
Private mObSymOff As Long
Private mObSymCnt As Long
Private mObStr As Long

'==============================================================================
' Form
'==============================================================================
Private Sub Form_Load()

    Dim active As String

    mExitCode = 1
    mLoading = True

    cbo_tool.AddItem "link"
    cbo_tool.AddItem "vbstatic"
    cbo_tool.AddItem "vbextern"

    LoadTools
    ParseCommandLine
    DetermineConfigFile
    RefreshCfgList

    active = IniGet(mCfgFile, "VBLinker", "Active", "")
    If Len(active) = 0 Or Not IniHas(mCfgFile, "Config." & active) Then
        If cbo_cfg.ListCount > 0 Then active = cbo_cfg.List(0) Else active = "Default"
    End If

    cbo_tool.ListIndex = 0
    LoadConfig active
    mLoading = False

    ShowStatus
    RefreshState

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'report the result to VB6 (the IDE must not be killed)
    If Not InIDE() Then ExitProcess mExitCode
End Sub

Private Function InIDE() As Boolean
    Debug.Assert MarkIDE(InIDE)
End Function

Private Function MarkIDE(ByRef b As Boolean) As Boolean
    b = True
    MarkIDE = True
End Function

Private Sub btn_exit_Click()
    mExitCode = 1
    Unload Me
End Sub

Private Sub opt_tab_Click(Index As Integer)
    fra_tab(0).Visible = (Index = 0)
    fra_tab(1).Visible = (Index = 1)
End Sub

Private Sub MarkDirty()
    If mLoading Then Exit Sub
    mDirty = True
    UpdateCaption
End Sub

Private Sub UpdateCaption()
    Me.Caption = "Linker - " & mCurCfg & IIf(mDirty, " *", "")
End Sub

Private Sub ShowStatus()
    lbl_status.Caption = "Configs: " & mCfgFile
    lbl_status.ToolTipText = mCfgFile
End Sub

'Enables / disables everything depending on the current options
Private Sub RefreshState()

    Dim isDll As Boolean
    Dim std As Boolean
    Dim s As Boolean
    Dim e As Boolean

    s = (check_static.Value = vbChecked)
    e = (check_extern.Value = vbChecked)

    lbl_static_obj.Enabled = s
    txt_static_obj.Enabled = s
    btn_static_auto.Enabled = s
    btn_static_browse.Enabled = s
    lbl_static_state.Enabled = s
    check_autofix.Enabled = s

    lbl_extern_obj.Enabled = e
    txt_extern_obj.Enabled = e
    btn_extern_auto.Enabled = e
    btn_extern_browse.Enabled = e
    lbl_extern_state.Enabled = e
    check_externall.Enabled = e

    UpdateArgsEnabled

    'DLL tab: only for /DLL builds, exports only for standard DLLs
    isDll = IsDllBuild()
    opt_tab(1).Enabled = isDll
    If Not isDll Then
        If opt_tab(1).Value Then opt_tab(0).Value = True
    End If

    std = isDll And opt_dll(1).Value
    opt_dll(0).Enabled = isDll
    opt_dll(1).Enabled = isDll
    lst_exports.Enabled = std
    btn_exp_add.Enabled = std
    btn_exp_addall.Enabled = std
    btn_exp_remove.Enabled = std
    btn_exp_clear.Enabled = std
    btn_exp_manual.Enabled = std
    lbl_exports.Enabled = std
    lbl_def.Enabled = std
    txt_def.Enabled = std
    btn_def_browse.Enabled = std
    btn_createdef.Enabled = std
    check_autodef.Enabled = std

End Sub

Private Sub UpdateArgsEnabled()

    Dim t As Long
    Dim en As Boolean

    t = cbo_tool.ListIndex
    If t < 0 Then Exit Sub

    Select Case t
        Case TOOL_STATIC: en = (check_static.Value = vbChecked)
        Case TOOL_EXTERN: en = (check_extern.Value = vbChecked)
        Case Else: en = True
    End Select

    args_list(t).Enabled = en
    btn_add.Enabled = en
    btn_remove.Enabled = en
    btn_addfile.Enabled = en
    btn_up.Enabled = en
    btn_down.Enabled = en

End Sub

Private Sub check_static_Click()
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub check_extern_Click()
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub check_autofix_Click()
    MarkDirty
End Sub

Private Sub check_externall_Click()
    MarkDirty
End Sub

Private Sub check_cmd_Click()
    MarkDirty
End Sub

Private Sub check_clean_Click()
    MarkDirty
End Sub

Private Sub check_autodef_Click()
    MarkDirty
End Sub

Private Sub opt_dll_Click(Index As Integer)
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub txt_def_Change()
    MarkDirty
End Sub

'==============================================================================
' Argument lists (0 = link, 1 = vbstatic, 2 = vbextern)
'==============================================================================
Private Sub cbo_tool_Click()

    Dim i As Long

    If cbo_tool.ListIndex < 0 Then Exit Sub
    For i = 0 To 2
        args_list(i).Visible = (i = cbo_tool.ListIndex)
    Next i
    UpdateArgsEnabled

End Sub

Private Sub AddArg(ByVal idx As Long, ByVal text As String, ByVal checked As Boolean, ByVal flag As Long)

    Dim old As Boolean

    old = mLoading
    mLoading = True
    args_list(idx).AddItem text
    args_list(idx).ItemData(args_list(idx).NewIndex) = flag
    args_list(idx).Selected(args_list(idx).NewIndex) = checked
    mLoading = old

End Sub

Private Sub args_list_ItemCheck(Index As Integer, Item As Integer)
    If mLoading Then Exit Sub
    MarkDirty
    If Index = TOOL_LINK Then RefreshState
End Sub

Private Sub args_list_DblClick(Index As Integer)

    Dim sel As Long
    Dim result As String
    Dim oldText As String
    Dim isOn As Boolean
    Dim flag As Long

    sel = args_list(Index).ListIndex
    If sel < 0 Then Exit Sub
    args_list(Index).Selected(sel) = Not args_list(Index).Selected(sel)

    oldText = args_list(Index).List(sel)
    result = Trim$(InputBox("Parameter:", , oldText))

    If LenB(result) > 0 And result <> oldText Then
        isOn = args_list(Index).Selected(sel)
        flag = args_list(Index).ItemData(sel)
        If flag = FLAG_CMD Then
            AddHidden oldText                      'the edited copy replaces the original
            flag = FLAG_USER
        End If
        mLoading = True
        args_list(Index).List(sel) = result
        args_list(Index).ItemData(sel) = flag
        args_list(Index).Selected(sel) = isOn
        mLoading = False
        MarkDirty
        RefreshState
    End If

End Sub

Private Sub btn_add_Click()

    Dim result As String

    result = Trim$(InputBox("Parameter:"))
    If LenB(result) > 0 Then
        AddArg cbo_tool.ListIndex, result, True, FLAG_USER
        MarkDirty
        RefreshState
    End If

End Sub

Private Sub btn_addfile_Click()

    Dim t As Long
    Dim files As Collection
    Dim v As Variant
    Dim p As String
    Dim flt As String

    t = cbo_tool.ListIndex
    If t < 0 Then Exit Sub

    flt = "All files (*.*)|*.*|Objects, libraries, resources|*.obj;*.lib;*.res;*.exp|DEF files (*.def)|*.def|DLL files (*.dll)|*.dll"
    Set files = FileDialog("Add file argument", flt, True, False, "")
    If files Is Nothing Then Exit Sub

    For Each v In files
        p = CStr(v)
        If t = TOOL_LINK And LCase$(Right$(p, 4)) = ".def" Then
            p = "/DEF:" & QIf(p)
        Else
            p = QIf(p)
        End If
        AddArg t, p, True, FLAG_USER
    Next v

    MarkDirty
    RefreshState

End Sub

Private Sub btn_remove_Click()

    Dim t As Long
    Dim sel As Long

    t = cbo_tool.ListIndex
    If t < 0 Then Exit Sub
    sel = args_list(t).ListIndex
    If sel < 0 Then Exit Sub

    If args_list(t).ItemData(sel) = FLAG_CMD Then AddHidden args_list(t).List(sel)

    mLoading = True
    args_list(t).RemoveItem sel
    If args_list(t).ListCount > 0 Then
        If sel >= args_list(t).ListCount Then sel = args_list(t).ListCount - 1
        args_list(t).ListIndex = sel
    End If
    mLoading = False

    MarkDirty
    RefreshState

End Sub

Private Sub btn_up_Click()
    MoveItem -1
End Sub

Private Sub btn_down_Click()
    MoveItem 1
End Sub

Private Sub MoveItem(ByVal stepv As Long)

    Dim t As Long
    Dim i As Long
    Dim j As Long
    Dim tx As String
    Dim ck As Boolean
    Dim fl As Long

    t = cbo_tool.ListIndex
    If t < 0 Then Exit Sub
    i = args_list(t).ListIndex
    j = i + stepv
    If i < 0 Or j < 0 Or j >= args_list(t).ListCount Then Exit Sub

    mLoading = True
    tx = args_list(t).List(i)
    ck = args_list(t).Selected(i)
    fl = args_list(t).ItemData(i)

    args_list(t).List(i) = args_list(t).List(j)
    args_list(t).ItemData(i) = args_list(t).ItemData(j)
    args_list(t).Selected(i) = args_list(t).Selected(j)

    args_list(t).List(j) = tx
    args_list(t).ItemData(j) = fl
    args_list(t).Selected(j) = ck

    args_list(t).ListIndex = j
    mLoading = False

    MarkDirty

End Sub

Private Sub AddHidden(ByVal text As String)

    Dim i As Long

    For i = 0 To mHiddenCount - 1
        If StrComp(mHidden(i), text, vbTextCompare) = 0 Then Exit Sub
    Next i
    ReDim Preserve mHidden(mHiddenCount)
    mHidden(mHiddenCount) = text
    mHiddenCount = mHiddenCount + 1

End Sub

Private Sub RemoveCmdArg(ByVal text As String)

    Dim i As Long

    For i = args_list(TOOL_LINK).ListCount - 1 To 0 Step -1
        If args_list(TOOL_LINK).ItemData(i) = FLAG_CMD Then
            If StrComp(args_list(TOOL_LINK).List(i), text, vbTextCompare) = 0 Then
                args_list(TOOL_LINK).RemoveItem i
            End If
        End If
    Next i

End Sub

Private Sub SetCmdArgOff(ByVal text As String)

    Dim i As Long

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).ItemData(i) = FLAG_CMD Then
            If StrComp(args_list(TOOL_LINK).List(i), text, vbTextCompare) = 0 Then
                args_list(TOOL_LINK).Selected(i) = False
            End If
        End If
    Next i

End Sub

'True if a checked link argument is /DLL
Private Function IsDllBuild() As Boolean

    Dim i As Long
    Dim u As String

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            u = UCase$(StripQ(args_list(TOOL_LINK).List(i)))
            If u = "/DLL" Or u = "-DLL" Then IsDllBuild = True: Exit Function
        End If
    Next i

End Function

'==============================================================================
' main.obj detection (COFF parsing)
'==============================================================================
Private Function U16(ByVal p As Long) As Long
    U16 = CLng(mOb(p)) + CLng(mOb(p + 1)) * 256&
End Function

Private Function I16(ByVal p As Long) As Long
    Dim v As Long
    v = U16(p)
    If v >= 32768 Then v = v - 65536
    I16 = v
End Function

Private Function I32(ByVal p As Long) As Long
    Dim v As Long
    v = CLng(mOb(p)) + CLng(mOb(p + 1)) * 256& + CLng(mOb(p + 2)) * 65536
    If mOb(p + 3) >= 128 Then
        v = v + (CLng(mOb(p + 3)) - 256&) * 16777216
    Else
        v = v + CLng(mOb(p + 3)) * 16777216
    End If
    I32 = v
End Function

Private Function ObjLoad(ByVal path As String) As Boolean

    Dim h As Long
    Dim n As Long

    On Error GoTo EH
    If Not FileExists(path) Then Exit Function
    n = FileLen(path)
    If n < 20 Then Exit Function

    ReDim mOb(n - 1)
    h = FreeFile
    Open path For Binary Access Read As #h
    Get #h, , mOb
    Close #h
    h = 0

    If U16(0) <> &H14C Then Exit Function              'not an i386 COFF object
    mObSymOff = I32(8)
    mObSymCnt = I32(12)
    If mObSymOff <= 0 Or mObSymCnt < 0 Then Exit Function
    mObStr = mObSymOff + 18& * mObSymCnt
    If mObStr + 4 > n Then Exit Function

    ObjLoad = True
    Exit Function
EH:
    On Error Resume Next
    If h <> 0 Then Close #h

End Function

Private Function ObjSymName(ByVal i As Long) As String

    Dim p As Long
    Dim k As Long
    Dim off As Long
    Dim s As String

    p = mObSymOff + i * 18&
    If I32(p) = 0 Then
        off = mObStr + I32(p + 4)
        Do While off <= UBound(mOb)
            If mOb(off) = 0 Then Exit Do
            s = s & Chr$(mOb(off))
            off = off + 1
        Loop
    Else
        For k = 0 To 7
            If mOb(p + k) = 0 Then Exit For
            s = s & Chr$(mOb(p + k))
        Next k
    End If
    ObjSymName = s

End Function

'Does the loaded object define the external symbol?
Private Function ObjDefines(ByVal symName As String) As Boolean

    Dim i As Long
    Dim p As Long

    On Error GoTo EH
    Do While i < mObSymCnt
        p = mObSymOff + i * 18&
        If p + 17 > UBound(mOb) Then Exit Do
        If mOb(p + 16) = 2 Then
            If I16(p + 12) > 0 Then
                If ObjSymName(i) = symName Then ObjDefines = True: Exit Function
            End If
        End If
        i = i + 1 + mOb(p + 17)
    Loop
EH:
End Function

'The project's main OBJ defines ___vbaS (the symbol VBStatic chains to)
Private Function IsMainObj(ByVal path As String) As Boolean
    If ObjLoad(path) Then IsMainObj = ObjDefines("___vbaS")
End Function

'Does the object contain VB6 Declare thunks (the 25 byte pattern VBExtern patches)?
Private Function ObjHasDeclares(ByVal path As String) As Boolean

    Dim p As Long
    Dim n As Long

    If Not ObjLoad(path) Then Exit Function
    n = UBound(mOb) - 24
    For p = 0 To n
        If mOb(p) = &HA1 Then
            If mOb(p + 5) = &HB And mOb(p + 6) = &HC0 And mOb(p + 7) = &H74 And mOb(p + 8) = 2 Then
                If mOb(p + 9) = &HFF And mOb(p + 10) = &HE0 And mOb(p + 11) = &H68 And mOb(p + 16) = &HB8 Then
                    If mOb(p + 21) = &HFF And mOb(p + 22) = &HD0 And mOb(p + 23) = &HFF And mOb(p + 24) = &HE0 Then
                        ObjHasDeclares = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next p

End Function

Private Function IsObjToken(ByVal token As String) As Boolean

    Dim p As String

    p = StripQ(token)
    If Len(p) < 5 Then Exit Function
    If Left$(p, 1) = "/" Or Left$(p, 1) = "-" Then Exit Function
    IsObjToken = (LCase$(Right$(p, 4)) = ".obj")

End Function

Private Function AddUnique(c As Collection, ByVal text As String) As Boolean
    On Error Resume Next
    c.Add text, LCase$(text)
    AddUnique = (Err.Number = 0)
    Err.Clear
End Function

'All checked .obj arguments of the link list
Private Function ObjsFromList() As Collection

    Dim c As New Collection
    Dim i As Long

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            If IsObjToken(args_list(TOOL_LINK).List(i)) Then
                AddUnique c, StripQ(args_list(TOOL_LINK).List(i))
            End If
        End If
    Next i
    Set ObjsFromList = c

End Function

Private Sub AutoDetectObjs()

    Dim objs As Collection
    Dim v As Variant
    Dim m As String
    Dim e As String
    Dim mHow As String
    Dim eHow As String

    Set objs = ObjsFromList()

    For Each v In objs
        If IsMainObj(CStr(v)) Then
            m = CStr(v)
            Exit For
        End If
    Next v

    If Len(m) > 0 Then
        mHow = "auto-detected (defines ___vbaS)"
    ElseIf objs.Count > 0 Then
        m = CStr(objs(1))
        mHow = "first .obj of the link list (___vbaS not found)"
    Else
        mHow = "no .obj in the link arguments - select it manually"
    End If
    txt_static_obj.Text = m
    lbl_static_state.Caption = mHow

    If Len(m) > 0 Then
        If ObjHasDeclares(m) Then
            e = m
            eHow = "auto-detected (contains Declare thunks)"
        End If
    End If
    If Len(e) = 0 Then
        For Each v In objs
            If ObjHasDeclares(CStr(v)) Then
                e = CStr(v)
                eHow = "auto-detected (first .obj with Declare thunks)"
                Exit For
            End If
        Next v
    End If
    If Len(e) = 0 Then
        e = m
        eHow = "main.obj (no Declare thunks found in the .obj files)"
    End If
    txt_extern_obj.Text = e
    lbl_extern_state.Caption = eHow

End Sub

Private Sub btn_static_auto_Click()
    AutoDetectObjs
End Sub

Private Sub btn_extern_auto_Click()
    AutoDetectObjs
End Sub

Private Sub btn_static_browse_Click()

    Dim c As Collection

    Set c = FileDialog("Select main.obj", "Object files (*.obj)|*.obj|All files (*.*)|*.*", False, False, txt_static_obj.Text)
    If c Is Nothing Then Exit Sub
    txt_static_obj.Text = c(1)
    lbl_static_state.Caption = "selected manually"

End Sub

Private Sub btn_extern_browse_Click()

    Dim c As Collection

    Set c = FileDialog("Select main.obj", "Object files (*.obj)|*.obj|All files (*.*)|*.*", False, False, txt_extern_obj.Text)
    If c Is Nothing Then Exit Sub
    txt_extern_obj.Text = c(1)
    lbl_extern_state.Caption = "selected manually"

End Sub

'==============================================================================
' DLL tab: function list, export list, .DEF
'==============================================================================
Private Sub ScanFunctions()

    Dim objs As Collection
    Dim seen As New Collection
    Dim v As Variant
    Dim i As Long
    Dim p As Long
    Dim sym As String
    Dim nm As String
    Dim sc As String

    mFnCount = 0
    Set objs = ObjsFromList()

    For Each v In objs
        If ObjLoad(CStr(v)) Then
            i = 0
            On Error Resume Next
            Do While i < mObSymCnt
                p = mObSymOff + i * 18&
                If p + 17 > UBound(mOb) Then Exit Do
                If mOb(p + 16) = 2 Then
                    If I16(p + 12) > 0 Then
                        sym = ObjSymName(i)
                        If Demangle(sym, nm, sc) Then
                            If AddUnique(seen, sym) Then
                                ReDim Preserve mFnName(mFnCount)
                                ReDim Preserve mFnScope(mFnCount)
                                ReDim Preserve mFnDecor(mFnCount)
                                mFnName(mFnCount) = nm
                                mFnScope(mFnCount) = sc
                                mFnDecor(mFnCount) = sym
                                mFnCount = mFnCount + 1
                            End If
                        End If
                    End If
                End If
                i = i + 1 + mOb(p + 17)
            Loop
            On Error GoTo 0
        End If
    Next v

    FillFuncList

End Sub

'?Name@Scope@@AAG...  ->  Name / Scope
Private Function Demangle(ByVal sym As String, ByRef nm As String, ByRef scope As String) As Boolean

    Dim p As Long
    Dim parts() As String

    nm = "": scope = ""
    If Left$(sym, 1) <> "?" Then Exit Function
    If Left$(sym, 2) = "??" Then Exit Function
    p = InStr(sym, "@@")
    If p < 3 Then Exit Function

    parts = Split(Mid$(sym, 2, p - 2), "@")
    nm = parts(0)
    If UBound(parts) >= 1 Then scope = parts(1)
    Demangle = (Len(nm) > 0)

End Function

Private Sub FillFuncList()

    Dim i As Long
    Dim f As String
    Dim t As String

    f = LCase$(Trim$(txt_filter.Text))
    lst_funcs.Clear
    For i = 0 To mFnCount - 1
        t = mFnName(i) & "  (" & mFnScope(i) & ")"
        If Len(f) = 0 Or InStr(LCase$(t), f) > 0 Then
            lst_funcs.AddItem t
            lst_funcs.ItemData(lst_funcs.NewIndex) = i
        End If
    Next i

End Sub

Private Sub txt_filter_Change()
    FillFuncList
End Sub

Private Sub btn_rescan_Click()
    ScanFunctions
End Sub

Private Function ExportNameOf(ByVal line As String) As String
    Dim p As Long
    p = InStr(line, "=")
    If p > 0 Then ExportNameOf = Trim$(Left$(line, p - 1)) Else ExportNameOf = Trim$(line)
End Function

Private Function ExportExists(ByVal decor As String, ByVal nm As String) As Boolean

    Dim i As Long
    Dim l As String

    For i = 0 To lst_exports.ListCount - 1
        l = lst_exports.List(i)
        If Len(decor) > 0 Then
            If Right$(l, Len(decor) + 1) = "=" & decor Then ExportExists = True: Exit Function
        End If
        If Len(nm) > 0 Then
            If StrComp(ExportNameOf(l), nm, vbTextCompare) = 0 Then ExportExists = True: Exit Function
        End If
    Next i

End Function

Private Sub AddExport(ByVal fi As Long)

    Dim nm As String

    If ExportExists(mFnDecor(fi), "") Then Exit Sub
    nm = mFnName(fi)
    If ExportExists("", nm) Then nm = mFnScope(fi) & "_" & mFnName(fi)   'name already taken
    lst_exports.AddItem nm & "=" & mFnDecor(fi)
    MarkDirty

End Sub

Private Sub btn_exp_add_Click()

    Dim i As Long

    For i = 0 To lst_funcs.ListCount - 1
        If lst_funcs.Selected(i) Then AddExport lst_funcs.ItemData(i)
    Next i

End Sub

Private Sub btn_exp_addall_Click()

    Dim i As Long

    For i = 0 To lst_funcs.ListCount - 1
        AddExport lst_funcs.ItemData(i)
    Next i

End Sub

Private Sub lst_funcs_DblClick()
    If lst_funcs.ListIndex >= 0 And lst_exports.Enabled Then AddExport lst_funcs.ItemData(lst_funcs.ListIndex)
End Sub

Private Sub btn_exp_remove_Click()

    Dim i As Long

    For i = lst_exports.ListCount - 1 To 0 Step -1
        If lst_exports.Selected(i) Then lst_exports.RemoveItem i
    Next i
    MarkDirty

End Sub

Private Sub btn_exp_clear_Click()
    lst_exports.Clear
    MarkDirty
End Sub

Private Sub btn_exp_manual_Click()

    Dim result As String

    result = Trim$(InputBox("Export (name, or name=internalName / name=?decorated@@name):", "Add export"))
    If LenB(result) > 0 Then
        If Not ExportExists("", ExportNameOf(result)) Then
            lst_exports.AddItem result
            MarkDirty
        End If
    End If

End Sub

Private Sub lst_exports_DblClick()

    Dim sel As Long
    Dim result As String

    sel = lst_exports.ListIndex
    If sel < 0 Then Exit Sub
    result = Trim$(InputBox("Export entry:", "Edit export", lst_exports.List(sel)))
    If LenB(result) > 0 Then
        lst_exports.List(sel) = result
        MarkDirty
    End If

End Sub

Private Function DefaultDefPath() As String

    Dim m As String
    Dim d As String

    m = Trim$(txt_static_obj.Text)
    If Len(m) = 0 Then m = Trim$(txt_extern_obj.Text)
    If Len(m) > 0 Then
        d = DirOf(m)
        If Len(d) = 0 Then d = mCfgDir
        DefaultDefPath = AddBS(d) & BaseOf(m) & ".def"
    Else
        DefaultDefPath = AddBS(mCfgDir) & "exports.def"
    End If

End Function

Private Sub btn_def_browse_Click()

    Dim c As Collection

    Set c = FileDialog("Select the .DEF file", "DEF files (*.def)|*.def|All files (*.*)|*.*", False, True, txt_def.Text)
    If c Is Nothing Then Exit Sub
    txt_def.Text = c(1)

End Sub

Private Sub btn_createdef_Click()
    If CreateDef(False) Then
        RefreshState
        MsgBox ".DEF created and added to the linker arguments:" & vbCrLf & Trim$(txt_def.Text), vbInformation
    End If
End Sub

Private Function CreateDef(ByVal quiet As Boolean) As Boolean

    Dim p As String
    Dim f As Long
    Dim i As Long

    p = Trim$(txt_def.Text)
    If Len(p) = 0 Then
        p = DefaultDefPath()
        txt_def.Text = p
    End If

    If lst_exports.ListCount = 0 Then
        If Not quiet Then MsgBox "The export list is empty.", vbExclamation
        Exit Function
    End If

    On Error GoTo EH
    f = FreeFile
    Open p For Output As #f
    Print #f, "; generated by Linker"
    Print #f, "EXPORTS"
    For i = 0 To lst_exports.ListCount - 1
        Print #f, "    " & lst_exports.List(i)
    Next i
    Close #f
    On Error GoTo 0

    SetLinkDef p
    CreateDef = True
    Exit Function

EH:
    On Error Resume Next
    Close #f
    MsgBox "Cannot write " & p & vbCrLf & Err.Description, vbExclamation

End Function

'Adds / replaces /DEF:<file> in the link arguments
Private Sub SetLinkDef(ByVal path As String)

    Dim i As Long
    Dim u As String
    Dim old As Boolean
    Dim txt As String

    txt = "/DEF:" & QIf(path)
    old = mLoading
    mLoading = True

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        u = UCase$(args_list(TOOL_LINK).List(i))
        If Left$(u, 5) = "/DEF:" Or Left$(u, 5) = "-DEF:" Then
            If args_list(TOOL_LINK).ItemData(i) = FLAG_CMD Then AddHidden args_list(TOOL_LINK).List(i)
            args_list(TOOL_LINK).List(i) = txt
            args_list(TOOL_LINK).ItemData(i) = FLAG_USER
            args_list(TOOL_LINK).Selected(i) = True
            mLoading = old
            MarkDirty
            Exit Sub
        End If
    Next i

    mLoading = old
    AddArg TOOL_LINK, txt, True, FLAG_USER
    MarkDirty

End Sub

'==============================================================================
' Build configs  (vblinker.ini in the directory of main.obj)
'==============================================================================
Private Sub ParseCommandLine()

    Dim a() As String
    Dim c As String

    mCmdCount = 0
    c = Trim$(Command$)
    If Len(c) = 0 Then Exit Sub

    a = SplitParams(Fix_Spaces(c))
    mCmdArgs = a
    mCmdCount = UBound(a) + 1

End Sub

Private Sub DetermineConfigFile()

    Dim i As Long
    Dim m As String
    Dim first As String
    Dim d As String

    For i = 0 To mCmdCount - 1
        If IsObjToken(mCmdArgs(i)) Then
            If Len(first) = 0 Then first = StripQ(mCmdArgs(i))
            If IsMainObj(StripQ(mCmdArgs(i))) Then
                m = StripQ(mCmdArgs(i))
                Exit For
            End If
        End If
    Next i
    If Len(m) = 0 Then m = first

    If Len(m) > 0 Then
        d = DirOf(m)
        If Len(d) = 0 Then d = CurDir$
        mCfgDir = d
    Else
        mCfgDir = App.Path
    End If
    mCfgFile = AddBS(mCfgDir) & "vblinker.ini"

End Sub

Private Function ListKey(ByVal k As Long) As String
    Select Case k
        Case TOOL_LINK: ListKey = "Link"
        Case TOOL_STATIC: ListKey = "Static"
        Case Else: ListKey = "Extern"
    End Select
End Function

Private Function BoolCheck(ByVal v As Long) As Long
    If v <> 0 Then BoolCheck = vbChecked Else BoolCheck = vbUnchecked
End Function

Private Function IniGet(ByVal f As String, ByVal sec As String, ByVal key As String, ByVal def As String) As String

    Dim buf As String
    Dim n As Long

    buf = String$(8192, 0)
    n = GetPrivateProfileString(sec, key, def, buf, Len(buf), f)
    IniGet = Left$(buf, n)

End Function

Private Function IniLong(ByVal f As String, ByVal sec As String, ByVal key As String, ByVal def As Long) As Long
    IniLong = CLng(Val(IniGet(f, sec, key, CStr(def))))
End Function

Private Sub IniPut(ByVal f As String, ByVal sec As String, ByVal key As String, ByVal txt As String)
    WritePrivateProfileString sec, key, txt, f
End Sub

Private Function IniHas(ByVal f As String, ByVal sec As String) As Boolean
    IniHas = (Len(IniGet(f, sec, "Version", "")) > 0)
End Function

'Values are stored with a leading "|" because GetPrivateProfileString strips surrounding quotes
Private Function Enc(ByVal s As String) As String
    Enc = "|" & s
End Function

Private Function Dec(ByVal s As String) As String
    If Left$(s, 1) = "|" Then Dec = Mid$(s, 2) Else Dec = s
End Function

Private Function CleanName(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim r As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Asc(ch) >= 32 And InStr("[]=;", ch) = 0 Then r = r & ch
    Next i
    CleanName = Left$(Trim$(r), 60)

End Function

Private Sub RefreshCfgList()

    Dim buf As String
    Dim n As Long
    Dim names() As String
    Dim i As Long
    Dim old As Boolean

    old = mLoading
    mLoading = True
    cbo_cfg.Clear
    buf = String$(32767, 0)
    n = GetPrivateProfileSectionNames(buf, Len(buf), mCfgFile)
    If n > 0 Then
        names = Split(Left$(buf, n), vbNullChar)
        For i = 0 To UBound(names)
            If LCase$(Left$(names(i), 7)) = "config." Then cbo_cfg.AddItem Mid$(names(i), 8)
        Next i
    End If
    mLoading = old

End Sub

Private Sub SelectCfg(ByVal cfgName As String)

    Dim i As Long
    Dim old As Boolean

    old = mLoading
    mLoading = True
    For i = 0 To cbo_cfg.ListCount - 1
        If StrComp(cbo_cfg.List(i), cfgName, vbTextCompare) = 0 Then
            cbo_cfg.ListIndex = i
            mLoading = old
            Exit Sub
        End If
    Next i
    cbo_cfg.AddItem cfgName
    cbo_cfg.ListIndex = cbo_cfg.NewIndex
    mLoading = old

End Sub

Private Sub LoadConfig(ByVal cfgName As String)

    Dim sec As String
    Dim f As String
    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim txt As String
    Dim kind As String
    Dim isOn As Boolean
    Dim old As Boolean

    old = mLoading
    mLoading = True
    sec = "Config." & cfgName
    f = mCfgFile

    'start from the command line of VB6
    args_list(TOOL_LINK).Clear
    args_list(TOOL_STATIC).Clear
    args_list(TOOL_EXTERN).Clear
    mHiddenCount = 0
    For i = 0 To mCmdCount - 1
        AddArg TOOL_LINK, mCmdArgs(i), True, FLAG_CMD
    Next i

    'defaults
    check_static.Value = vbUnchecked
    check_extern.Value = vbUnchecked
    check_externall.Value = vbUnchecked
    check_autofix.Value = vbChecked
    check_clean.Value = vbChecked
    check_cmd.Value = vbUnchecked
    check_autodef.Value = vbChecked
    opt_dll(0).Value = True
    lst_exports.Clear
    txt_def.Text = ""

    If IniHas(f, sec) Then

        check_static.Value = BoolCheck(IniLong(f, sec, "Static", 0))
        check_extern.Value = BoolCheck(IniLong(f, sec, "Extern", 0))
        check_externall.Value = BoolCheck(IniLong(f, sec, "ExternAll", 0))
        check_autofix.Value = BoolCheck(IniLong(f, sec, "AutoFix", 1))
        check_clean.Value = BoolCheck(IniLong(f, sec, "Clean", 1))
        check_cmd.Value = BoolCheck(IniLong(f, sec, "RunInCmd", 0))
        check_autodef.Value = BoolCheck(IniLong(f, sec, "AutoDef", 1))
        If IniLong(f, sec, "DllMode", 0) = 1 Then opt_dll(1).Value = True Else opt_dll(0).Value = True
        txt_def.Text = IniGet(f, sec, "DefFile", "")

        n = IniLong(f, sec, "Hide.Count", 0)
        For i = 0 To n - 1
            txt = Dec(IniGet(f, sec, "Hide." & i, ""))
            If Len(txt) > 0 Then
                AddHidden txt
                RemoveCmdArg txt
            End If
        Next i

        For k = 0 To 2
            n = IniLong(f, sec, ListKey(k) & ".Count", 0)
            For i = 0 To n - 1
                txt = Dec(IniGet(f, sec, ListKey(k) & "." & i, ""))
                kind = IniGet(f, sec, ListKey(k) & "." & i & ".Kind", "user")
                isOn = (IniLong(f, sec, ListKey(k) & "." & i & ".On", 1) <> 0)
                If Len(txt) > 0 Then
                    If kind = "off" Then
                        SetCmdArgOff txt
                    Else
                        AddArg k, txt, isOn, FLAG_USER
                    End If
                End If
            Next i
        Next k

        n = IniLong(f, sec, "Exports.Count", 0)
        For i = 0 To n - 1
            txt = Dec(IniGet(f, sec, "Exports." & i, ""))
            If Len(txt) > 0 Then lst_exports.AddItem txt
        Next i

    Else
        'fresh config: handy (disabled) examples for VBExtern
        AddArg TOOL_EXTERN, """/PLACEHOLDER:<extern>""", False, FLAG_USER
        AddArg TOOL_EXTERN, "/PREFIX:_", False, FLAG_USER
    End If

    mCurCfg = cfgName
    SelectCfg cfgName
    AutoDetectObjs
    If Len(Trim$(txt_def.Text)) = 0 Then txt_def.Text = DefaultDefPath()
    ScanFunctions

    mLoading = old
    mDirty = False
    UpdateCaption
    RefreshState

End Sub

Private Function SaveConfig(ByVal cfgName As String) As Boolean

    Dim sec As String
    Dim f As String
    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim kind As String

    sec = "Config." & cfgName
    f = mCfgFile

    WritePrivateProfileString sec, vbNullString, vbNullString, f           'drop the old section
    If WritePrivateProfileString(sec, "Version", "1", f) = 0 Then
        MsgBox "Cannot write the config file:" & vbCrLf & f, vbExclamation
        Exit Function
    End If

    IniPut f, sec, "Static", IIf(check_static.Value = vbChecked, "1", "0")
    IniPut f, sec, "Extern", IIf(check_extern.Value = vbChecked, "1", "0")
    IniPut f, sec, "ExternAll", IIf(check_externall.Value = vbChecked, "1", "0")
    IniPut f, sec, "AutoFix", IIf(check_autofix.Value = vbChecked, "1", "0")
    IniPut f, sec, "Clean", IIf(check_clean.Value = vbChecked, "1", "0")
    IniPut f, sec, "RunInCmd", IIf(check_cmd.Value = vbChecked, "1", "0")
    IniPut f, sec, "AutoDef", IIf(check_autodef.Value = vbChecked, "1", "0")
    IniPut f, sec, "DllMode", IIf(opt_dll(1).Value, "1", "0")
    IniPut f, sec, "DefFile", Trim$(txt_def.Text)

    'argument lists: user entries in full, command line entries only when unchecked
    For k = 0 To 2
        n = 0
        For i = 0 To args_list(k).ListCount - 1
            kind = ""
            If args_list(k).ItemData(i) = FLAG_CMD Then
                If Not args_list(k).Selected(i) Then kind = "off"
            Else
                kind = "user"
            End If
            If Len(kind) > 0 Then
                IniPut f, sec, ListKey(k) & "." & n, Enc(args_list(k).List(i))
                IniPut f, sec, ListKey(k) & "." & n & ".Kind", kind
                IniPut f, sec, ListKey(k) & "." & n & ".On", IIf(args_list(k).Selected(i), "1", "0")
                n = n + 1
            End If
        Next i
        IniPut f, sec, ListKey(k) & ".Count", CStr(n)
    Next k

    IniPut f, sec, "Hide.Count", CStr(mHiddenCount)
    For i = 0 To mHiddenCount - 1
        IniPut f, sec, "Hide." & i, Enc(mHidden(i))
    Next i

    IniPut f, sec, "Exports.Count", CStr(lst_exports.ListCount)
    For i = 0 To lst_exports.ListCount - 1
        IniPut f, sec, "Exports." & i, Enc(lst_exports.List(i))
    Next i

    IniPut f, "VBLinker", "Active", cfgName
    SaveConfig = True

End Function

Private Sub cbo_cfg_Click()

    Dim nm As String

    If mLoading Then Exit Sub
    If cbo_cfg.ListIndex < 0 Then Exit Sub
    nm = cbo_cfg.List(cbo_cfg.ListIndex)
    If StrComp(nm, mCurCfg, vbTextCompare) = 0 Then Exit Sub

    If mDirty Then
        If MsgBox("Discard the unsaved changes of '" & mCurCfg & "'?", vbYesNo + vbQuestion) <> vbYes Then
            SelectCfg mCurCfg
            Exit Sub
        End If
    End If

    LoadConfig nm
    IniPut mCfgFile, "VBLinker", "Active", nm

End Sub

Private Sub btn_cfgsave_Click()
    If Len(mCurCfg) = 0 Then mCurCfg = "Default"
    If SaveConfig(mCurCfg) Then
        mDirty = False
        UpdateCaption
    End If
End Sub

Private Sub btn_cfgsaveas_Click()

    Dim nm As String

    nm = CleanName(InputBox("Name of the build config:", "Save as", mCurCfg))
    If Len(nm) = 0 Then Exit Sub

    If StrComp(nm, mCurCfg, vbTextCompare) <> 0 Then
        If IniHas(mCfgFile, "Config." & nm) Then
            If MsgBox("Overwrite the build config '" & nm & "'?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        End If
    End If

    If SaveConfig(nm) Then
        mCurCfg = nm
        RefreshCfgList
        SelectCfg nm
        mDirty = False
        UpdateCaption
    End If

End Sub

Private Sub btn_cfgdel_Click()

    Dim nm As String

    If MsgBox("Delete the build config '" & mCurCfg & "'?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    WritePrivateProfileString "Config." & mCurCfg, vbNullString, vbNullString, mCfgFile

    RefreshCfgList
    If cbo_cfg.ListCount > 0 Then nm = cbo_cfg.List(0) Else nm = "Default"
    LoadConfig nm
    IniPut mCfgFile, "VBLinker", "Active", nm

End Sub

'==============================================================================
' Tool paths (global, next to this exe)
'==============================================================================
Private Function AppIni() As String
    AppIni = AddBS(App.Path) & "vblinker.ini"
End Function

Private Function DefaultRuntime() As String

    Dim p As String

    p = Environ$("windir") & "\SysWOW64\msvbvm60.dll"
    If Not FileExists(p) Then p = Environ$("windir") & "\System32\msvbvm60.dll"
    DefaultRuntime = p

End Function

Private Sub LoadTools()
    mToolLink = IniGet(AppIni(), "Tools", "Link2", AddBS(App.Path) & "LINK2.EXE")
    mToolStatic = IniGet(AppIni(), "Tools", "VBStatic", AddBS(App.Path) & "VBStatic.exe")
    mToolExtern = IniGet(AppIni(), "Tools", "VBExtern", AddBS(App.Path) & "VBExtern.exe")
    mRuntime = IniGet(AppIni(), "Tools", "Runtime", DefaultRuntime())
End Sub

Private Sub SaveTools()
    IniPut AppIni(), "Tools", "Link2", mToolLink
    IniPut AppIni(), "Tools", "VBStatic", mToolStatic
    IniPut AppIni(), "Tools", "VBExtern", mToolExtern
    IniPut AppIni(), "Tools", "Runtime", mRuntime
End Sub

Private Sub btn_tools_Click()
    mnu_tool_link.Caption = "LINK2.EXE: " & Replace(mToolLink, "&", "&&")
    mnu_tool_static.Caption = "VBStatic.exe: " & Replace(mToolStatic, "&", "&&")
    mnu_tool_extern.Caption = "VBExtern.exe: " & Replace(mToolExtern, "&", "&&")
    mnu_tool_runtime.Caption = "MSVBVM60.DLL: " & Replace(mRuntime, "&", "&&")
    PopupMenu mnu_tools
End Sub

Private Sub PickTool(ByRef target As String, ByVal title As String, ByVal flt As String)

    Dim c As Collection

    Set c = FileDialog(title, flt, False, False, target)
    If c Is Nothing Then Exit Sub
    target = c(1)
    SaveTools

End Sub

Private Sub mnu_tool_link_Click()
    PickTool mToolLink, "Select the real linker (LINK2.EXE)", "Executables (*.exe)|*.exe|All files (*.*)|*.*"
End Sub

Private Sub mnu_tool_static_Click()
    PickTool mToolStatic, "Select VBStatic.exe", "Executables (*.exe)|*.exe|All files (*.*)|*.*"
End Sub

Private Sub mnu_tool_extern_Click()
    PickTool mToolExtern, "Select VBExtern.exe", "Executables (*.exe)|*.exe|All files (*.*)|*.*"
End Sub

Private Sub mnu_tool_runtime_Click()
    PickTool mRuntime, "Select MSVBVM60.DLL", "DLL files (*.dll)|*.dll|All files (*.*)|*.*"
End Sub

'==============================================================================
' Build
'==============================================================================
Private Sub btn_preview_Click()

    Dim rep As String

    RunPipeline True, rep
    Clipboard.Clear
    Clipboard.SetText rep
    MsgBox "The commands were copied to the clipboard:" & vbCrLf & vbCrLf & Left$(rep, 1500), vbInformation

End Sub

Private Sub btn_ok_Click()

    Dim rc As Long
    Dim rep As String

    If Not CheckSetup() Then Exit Sub

    If IsDllBuild() And opt_dll(1).Value And check_autodef.Value = vbChecked And lst_exports.ListCount > 0 Then
        If Not CreateDef(True) Then Exit Sub
    End If

    Me.Enabled = False
    rc = RunPipeline(False, rep)
    Me.Enabled = True

    If rc = 0 Then
        mExitCode = 0
        Unload Me
    Else
        mExitCode = rc
    End If

End Sub

Private Function CheckSetup() As Boolean

    Dim s As Boolean
    Dim e As Boolean
    Dim isDll As Boolean

    s = (check_static.Value = vbChecked)
    e = (check_extern.Value = vbChecked)
    isDll = IsDllBuild()

    If Not FileExists(mToolLink) Then
        MsgBox "Linker not found:" & vbCrLf & mToolLink & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
        Exit Function
    End If

    If e Then
        If Not FileExists(mToolExtern) Then
            MsgBox "VBExtern.exe not found:" & vbCrLf & mToolExtern & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If check_externall.Value <> vbChecked Then
            If Not FileExists(Trim$(txt_extern_obj.Text)) Then
                MsgBox "VBExtern: main.obj not found:" & vbCrLf & txt_extern_obj.Text, vbExclamation
                Exit Function
            End If
        End If
    End If

    If s Then
        If Not FileExists(mToolStatic) Then
            MsgBox "VBStatic.exe not found:" & vbCrLf & mToolStatic & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If Not FileExists(mRuntime) Then
            MsgBox "MSVBVM60.DLL not found:" & vbCrLf & mRuntime & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If Not FileExists(Trim$(txt_static_obj.Text)) Then
            MsgBox "VBStatic: main.obj not found:" & vbCrLf & txt_static_obj.Text, vbExclamation
            Exit Function
        End If
    End If

    If isDll And opt_dll(1).Value Then
        If Not s Then
            If MsgBox("Standard DLL without VBStatic: the ActiveX exports of the VB6 object stay in place." & vbCrLf & _
                      "Continue anyway?", vbYesNo + vbQuestion) <> vbYes Then Exit Function
        ElseIf lst_exports.ListCount = 0 Then
            If MsgBox("The export list is empty - the DLL will have no exports." & vbCrLf & _
                      "Continue anyway?", vbYesNo + vbQuestion) <> vbYes Then Exit Function
        End If
    End If

    CheckSetup = True

End Function

'Runs (or, with dryRun, only lists) VBExtern -> VBStatic -> LINK2. Returns 0 on success.
Private Function RunPipeline(ByVal dryRun As Boolean, ByRef report As String) As Long

    Dim isDll As Boolean
    Dim useStatic As Boolean
    Dim useExtern As Boolean
    Dim targets As Collection
    Dim objs As Collection
    Dim v As Variant
    Dim src As String
    Dim dst As String
    Dim cmd As String
    Dim rc As Long
    Dim result As Long

    isDll = IsDllBuild()
    useStatic = (check_static.Value = vbChecked)
    useExtern = (check_extern.Value = vbChecked)
    mMapCount = 0
    mTempCount = 0
    report = ""

    '---- 1. VBExtern ---------------------------------------------------------
    If useExtern Then
        Set targets = New Collection
        If check_externall.Value = vbChecked Then
            Set objs = ObjsFromList()
            For Each v In objs
                If ObjHasDeclares(CStr(v)) Then targets.Add CStr(v)
            Next v
        ElseIf Len(Trim$(txt_extern_obj.Text)) > 0 Then
            targets.Add Trim$(txt_extern_obj.Text)
        End If

        For Each v In targets
            src = CStr(v)
            dst = DerivePath(src, ".vbx")
            cmd = Q(mToolExtern) & ArgsFor(TOOL_EXTERN, True) & " " & Q(src) & " /OUT:" & Q(dst) & ArgsFor(TOOL_EXTERN, False)
            If dryRun Then
                report = report & cmd & vbCrLf
                MapSet src, dst
            Else
                DelFile dst
                rc = RunStep("VBExtern - " & FileOf(src), cmd, False)
                If rc <> 0 Then
                    StepFailed "VBExtern", rc
                    result = rc
                    GoTo Done
                End If
                'VBExtern only writes /OUT: when it really patched something
                If FileExists(dst) Then
                    MapSet src, dst
                    AddTemp dst
                End If
            End If
        Next v
    End If

    '---- 2. VBStatic ---------------------------------------------------------
    If useStatic Then
        src = Trim$(txt_static_obj.Text)
        dst = DerivePath(src, ".vbs")
        cmd = Q(mToolStatic) & ArgsFor(TOOL_STATIC, True)
        If isDll Then
            If opt_dll(1).Value Then cmd = cmd & " /DLL:STANDARD" Else cmd = cmd & " /DLL:ACTIVEX"
        End If
        cmd = cmd & " /OUT:" & Q(dst) & " " & Q(mRuntime) & " " & Q(MapGet(src)) & ArgsFor(TOOL_STATIC, False)
        If dryRun Then
            report = report & cmd & vbCrLf
            MapSet src, dst
        Else
            DelFile dst
            rc = RunStep("VBStatic - " & FileOf(src), cmd, False)
            'VBStatic always exits with 0, success = the output file exists
            If rc <> 0 Or Not FileExists(dst) Then
                If rc = 0 Then rc = 1
                StepFailed "VBStatic", rc
                result = rc
                GoTo Done
            End If
            MapSet src, dst
            AddTemp dst
        End If
    End If

    '---- 3. link -------------------------------------------------------------
    cmd = Q(mToolLink) & BuildLinkArgs(isDll, useStatic)
    If dryRun Then
        report = report & cmd & vbCrLf
    Else
        rc = RunStep("LINK2", cmd, True)
        If rc <> 0 Then
            StepFailed "LINK2", rc
            result = rc
        End If
    End If

Done:
    If result = 0 And Not dryRun And check_clean.Value = vbChecked Then CleanTemps
    RunPipeline = result

End Function

Private Sub StepFailed(ByVal tool As String, ByVal rc As Long)

    Dim msg As String

    If rc = -1 Then
        msg = "Could not start " & tool & "."
    Else
        msg = tool & " failed (exit code " & rc & ")."
    End If
    If check_cmd.Value <> vbChecked Then msg = msg & vbCrLf & "Enable 'Run in CMD' to see the output of the tool."
    MsgBox msg, vbExclamation

End Sub

'Switches (/xxx) or positional arguments of the vbstatic / vbextern list
Private Function ArgsFor(ByVal tool As Long, ByVal switches As Boolean) As String

    Dim i As Long
    Dim t As String
    Dim c As String
    Dim s As String

    For i = 0 To args_list(tool).ListCount - 1
        If args_list(tool).Selected(i) Then
            t = args_list(tool).List(i)
            c = Left$(StripQ(t), 1)
            If (c = "/" Or c = "-") = switches Then s = s & " " & t
        End If
    Next i
    ArgsFor = s

End Function

Private Function BuildLinkArgs(ByVal isDll As Boolean, ByVal useStatic As Boolean) As String

    Dim i As Long
    Dim t As String
    Dim u As String
    Dim p As String
    Dim s As String
    Dim entry As String
    Dim doFix As Boolean
    Dim skip As Boolean
    Dim hasEntry As Boolean
    Dim hasNoDef As Boolean

    doFix = useStatic And (check_autofix.Value = vbChecked)
    If isDll Then entry = "_VBStaticDllEntry@12" Else entry = "_VBStaticEntry"

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            t = args_list(TOOL_LINK).List(i)
            p = StripQ(t)
            u = UCase$(p)
            skip = False

            If IsObjToken(t) Then
                If MapGet(p) <> p Then t = Q(MapGet(p))          'patched copy
            ElseIf doFix Then
                If InStr(u, "MSVBVM60") > 0 And Right$(u, 4) = ".LIB" Then
                    skip = True                                  'runtime is embedded now
                ElseIf Left$(u, 7) = "/ENTRY:" Or Left$(u, 7) = "-ENTRY:" Then
                    t = "/ENTRY:" & entry
                    hasEntry = True
                ElseIf u = "/NODEFAULTLIB" Or u = "-NODEFAULTLIB" Then
                    hasNoDef = True
                End If
            End If

            If Not skip Then s = s & " " & t
        End If
    Next i

    If doFix Then
        If Not hasEntry Then s = s & " /ENTRY:" & entry
        If Not hasNoDef Then s = s & " /NODEFAULTLIB"
    End If
    BuildLinkArgs = s

End Function

Private Sub MapSet(ByVal orig As String, ByVal patched As String)

    Dim i As Long

    For i = 0 To mMapCount - 1
        If StrComp(mMapOrig(i), orig, vbTextCompare) = 0 Then
            mMapNew(i) = patched
            Exit Sub
        End If
    Next i
    ReDim Preserve mMapOrig(mMapCount)
    ReDim Preserve mMapNew(mMapCount)
    mMapOrig(mMapCount) = orig
    mMapNew(mMapCount) = patched
    mMapCount = mMapCount + 1

End Sub

Private Function MapGet(ByVal orig As String) As String

    Dim i As Long

    For i = 0 To mMapCount - 1
        If StrComp(mMapOrig(i), orig, vbTextCompare) = 0 Then
            MapGet = mMapNew(i)
            Exit Function
        End If
    Next i
    MapGet = orig

End Function

Private Sub AddTemp(ByVal path As String)
    ReDim Preserve mTemp(mTempCount)
    mTemp(mTempCount) = path
    mTempCount = mTempCount + 1
End Sub

Private Sub CleanTemps()

    Dim i As Long

    For i = 0 To mTempCount - 1
        DelFile mTemp(i)
    Next i
    mTempCount = 0

End Sub

Private Sub DelFile(ByVal path As String)
    On Error Resume Next
    If FileExists(path) Then Kill path
End Sub

Private Function DerivePath(ByVal src As String, ByVal suffix As String) As String

    Dim d As String

    d = DirOf(src)
    If Len(d) > 0 Then
        DerivePath = d & "\" & BaseOf(src) & suffix & ".obj"
    Else
        DerivePath = BaseOf(src) & suffix & ".obj"
    End If

End Function

'Runs one step and waits for it. Returns the exit code (-1 = could not start).
Private Function RunStep(ByVal title As String, ByVal cmdline As String, ByVal isLast As Boolean) As Long

    Dim tmp As String
    Dim f As Long
    Dim qq As String

    If check_cmd.Value <> vbChecked Then
        RunStep = RunWait(cmdline, SW_SHOWMINNOACTIVE)
        Exit Function
    End If

    'visible console: a small batch file keeps the window open and returns the exit code
    tmp = Environ$("TEMP")
    If Len(tmp) = 0 Then tmp = App.Path
    tmp = AddBS(tmp) & "vblinker_step.cmd"
    qq = Chr$(34)

    f = FreeFile
    Open tmp For Output As #f
    Print #f, "@echo off"
    Print #f, "title " & CmdEscape(title)
    Print #f, "echo ==== " & CmdEscape(title) & " ===="
    Print #f, "echo."
    Print #f, "start " & qq & qq & " /wait /b " & Replace(cmdline, "%", "%%")
    Print #f, "set rc=%errorlevel%"
    Print #f, "echo."
    Print #f, "echo [exit code %rc%]"
    If isLast Then
        Print #f, "pause"
    Else
        Print #f, "if not " & qq & "%rc%" & qq & "==" & qq & "0" & qq & " pause"
    End If
    Print #f, "exit /b %rc%"
    Close #f

    RunStep = RunWait(Environ$("ComSpec") & " /c " & qq & qq & tmp & qq & qq, SW_SHOWNORMAL)
    DelFile tmp

End Function

Private Function RunWait(ByVal cmdline As String, ByVal showMode As Long) As Long

    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    Dim rc As Long
    Dim code As Long

    si.cb = Len(si)
    si.dwFlags = STARTF_USESHOWWINDOW
    si.wShowWindow = showMode

    rc = CreateProcess(vbNullString, cmdline, ByVal 0&, ByVal 0&, 0&, 0&, ByVal 0&, vbNullString, si, pi)
    If rc = 0 Then
        RunWait = -1
        Exit Function
    End If

    Do
        rc = WaitForSingleObject(pi.hProcess, 50)
        DoEvents
    Loop While rc = WAIT_TIMEOUT

    GetExitCodeProcess pi.hProcess, code
    CloseHandle pi.hThread
    CloseHandle pi.hProcess
    RunWait = code

End Function

Private Function CmdEscape(ByVal s As String) As String
    s = Replace(s, "^", "^^")
    s = Replace(s, "&", "^&")
    s = Replace(s, "<", "^<")
    s = Replace(s, ">", "^>")
    s = Replace(s, "|", "^|")
    s = Replace(s, "(", "^(")
    s = Replace(s, ")", "^)")
    s = Replace(s, "%", "%%")
    CmdEscape = s
End Function

'==============================================================================
' File dialog (comdlg32, no OCX needed). Returns the selected paths or Nothing.
' flt = "Description|*.ext|Description|*.ext"
'==============================================================================
Private Function FileDialog(ByVal title As String, ByVal flt As String, ByVal multi As Boolean, ByVal forSave As Boolean, ByVal initPath As String) As Collection

    Dim ofn As OPENFILENAME
    Dim buf As String
    Dim s As String
    Dim p As Long
    Dim parts() As String
    Dim i As Long
    Dim res As Collection
    Dim d As String

    initPath = StripQ(initPath)
    buf = initPath & String$(4096 - Len(initPath), 0)

    d = DirOf(initPath)
    If Len(d) = 0 Then d = mCfgDir

    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.lpstrFilter = Replace(flt, "|", vbNullChar) & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = buf
    ofn.nMaxFile = Len(buf)
    ofn.lpstrInitialDir = d
    ofn.lpstrTitle = title
    ofn.flags = OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_PATHMUSTEXIST
    If forSave Then
        ofn.flags = ofn.flags Or OFN_OVERWRITEPROMPT
    Else
        ofn.flags = ofn.flags Or OFN_FILEMUSTEXIST
    End If
    If multi Then ofn.flags = ofn.flags Or OFN_ALLOWMULTISELECT

    If forSave Then
        p = GetSaveFileName(ofn)
    Else
        p = GetOpenFileName(ofn)
    End If
    If p = 0 Then Exit Function

    s = ofn.lpstrFile
    p = InStr(s, vbNullChar & vbNullChar)
    If p > 0 Then
        s = Left$(s, p - 1)
    Else
        p = InStr(s, vbNullChar)
        If p > 0 Then s = Left$(s, p - 1)
    End If
    If Len(s) = 0 Then Exit Function

    Set res = New Collection
    parts = Split(s, vbNullChar)
    If UBound(parts) = 0 Then
        res.Add parts(0)                                    'one file: full path
    Else
        For i = 1 To UBound(parts)                          'several: folder, file, file ...
            res.Add AddBS(parts(0)) & parts(i)
        Next i
    End If
    Set FileDialog = res

End Function

'==============================================================================
' Small helpers
'==============================================================================
Private Function Q(ByVal s As String) As String
    Q = """" & s & """"
End Function

Private Function QIf(ByVal s As String) As String
    If InStr(s, " ") > 0 Then QIf = """" & s & """" Else QIf = s
End Function

Private Function StripQ(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) >= 2 Then
        If Left$(s, 1) = """" And Right$(s, 1) = """" Then s = Mid$(s, 2, Len(s) - 2)
    End If
    StripQ = s
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    FileExists = ((GetAttr(path) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
    Err.Clear
End Function

Private Function AddBS(ByVal d As String) As String
    If Right$(d, 1) = "\" Then AddBS = d Else AddBS = d & "\"
End Function

Private Function DirOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 0 Then DirOf = Left$(p, i - 1)
End Function

Private Function FileOf(ByVal p As String) As String
    FileOf = Mid$(p, InStrRev(p, "\") + 1)
End Function

Private Function BaseOf(ByVal p As String) As String
    Dim f As String
    Dim i As Long
    f = FileOf(p)
    i = InStrRev(f, ".")
    If i > 1 Then BaseOf = Left$(f, i - 1) Else BaseOf = f
End Function

'==============================================================================
' Command line splitting (unchanged from the original form)
'==============================================================================
Public Function SplitParams(ByVal Text As String) As String()
    Dim result() As String
    Dim current As String
    Dim i As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim count As Long
    
    ReDim result(0)
    
    For i = 1 To Len(Text)
        ch = Mid$(Text, i, 1)
        
        Select Case ch
            
            Case """"
                inQuotes = Not inQuotes
                current = current & ch   ' <<< KEEP the quote
                
            Case " "
                If inQuotes Then
                    current = current & ch
                Else
                    If Len(current) > 0 Then
                        result(count) = current
                        count = count + 1
                        ReDim Preserve result(0 To count)
                        current = ""
                    End If
                End If
                
            Case Else
                current = current & ch
                
        End Select
    Next i
    
    ' Add final token
    If Len(current) > 0 Then
        result(count) = current
    ElseIf count > 0 Then
        ReDim Preserve result(0 To count - 1)
    End If
    
    SplitParams = result
End Function

Public Function Fix_Spaces(ByVal str As String) As String
    
    Dim i As Long
    Dim j As Long
    Dim isStr As Boolean
    
    Const SPACE_CHAR = " "
    
    Fix_Spaces = Space$(Len(str))
    j = 1
    
    For i = 1 To Len(str)
        If (Mid$(str, i, 1) = SPACE_CHAR) And Not isStr Then
            If Mid$(str, i + 1, 1) <> SPACE_CHAR Then
                Mid$(Fix_Spaces, j, 1) = SPACE_CHAR
                j = j + 1
            End If
        ElseIf Mid$(str, i, 1) = """" Then
            isStr = Not isStr
        
            Mid$(Fix_Spaces, j, 1) = Mid$(str, i, 1)
            j = j + 1
        Else
            Mid$(Fix_Spaces, j, 1) = Mid$(str, i, 1)
            j = j + 1
        End If
    Next i
    
    Fix_Spaces = RTrim$(Fix_Spaces)
    
End Function

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Linker"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8040
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cbo_cfg 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown-Liste
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
      Style           =   1  'Grafisch
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
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   600
      Width           =   1500
   End
   Begin VB.OptionButton opt_tab 
      Caption         =   "Icon"
      Height          =   345
      Index           =   2
      Left            =   3120
      Style           =   1  'Grafisch
      TabIndex        =   8
      Top             =   600
      Width           =   1500
   End
   Begin VB.OptionButton opt_tab 
      Caption         =   "Info"
      Height          =   345
      Index           =   3
      Left            =   4620
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   600
      Width           =   1500
   End
   Begin VB.Frame fra_tab 
      Height          =   4680
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   7800
      Begin VB.CheckBox check_static 
         Caption         =   "Embed MSVBVM60.DLL into build (VBStatic)"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   7400
      End
      Begin VB.TextBox txt_static_obj 
         Height          =   315
         Left            =   1300
         TabIndex        =   29
         Top             =   420
         Width           =   4300
      End
      Begin VB.CommandButton btn_static_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5700
         TabIndex        =   30
         Top             =   420
         Width           =   600
      End
      Begin VB.CommandButton btn_static_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   6350
         TabIndex        =   31
         Top             =   420
         Width           =   450
      End
      Begin VB.CheckBox check_autofix 
         Caption         =   "Adjust link args for VBStatic (entry point, drop msvbvm60.lib, /NODEFAULTLIB)"
         Height          =   240
         Left            =   480
         TabIndex        =   33
         Top             =   1000
         Value           =   1  'Aktiviert
         Width           =   7200
      End
      Begin VB.CheckBox check_extern 
         Caption         =   "Support for external symbols (VBExtern)"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1380
         Width           =   7400
      End
      Begin VB.TextBox txt_extern_obj 
         Height          =   315
         Left            =   1300
         TabIndex        =   36
         Top             =   1680
         Width           =   4300
      End
      Begin VB.CommandButton btn_extern_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5700
         TabIndex        =   37
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
      Begin VB.ComboBox cbo_tool 
         Height          =   315
         Left            =   1450
         Style           =   2  'Dropdown-Liste
         TabIndex        =   14
         Top             =   2520
         Width           =   2400
      End
      Begin VB.ListBox args_list 
         Height          =   1410
         Index           =   0
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   15
         Top             =   2940
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1410
         Index           =   1
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   16
         Top             =   2940
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1410
         Index           =   2
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   17
         Top             =   2940
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1410
         Index           =   3
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   18
         Top             =   2940
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.CommandButton btn_add 
         Caption         =   "+"
         Height          =   285
         Left            =   7000
         TabIndex        =   19
         Top             =   2940
         Width           =   720
      End
      Begin VB.CommandButton btn_remove 
         Caption         =   "-"
         Height          =   285
         Left            =   7000
         TabIndex        =   20
         Top             =   3240
         Width           =   720
      End
      Begin VB.CommandButton btn_addfile 
         Caption         =   "File..."
         Height          =   285
         Left            =   7000
         TabIndex        =   21
         Top             =   3540
         Width           =   720
      End
      Begin VB.CommandButton btn_up 
         Caption         =   "Up"
         Height          =   285
         Left            =   7000
         TabIndex        =   22
         Top             =   3840
         Width           =   720
      End
      Begin VB.CommandButton btn_down 
         Caption         =   "Down"
         Height          =   285
         Left            =   7000
         TabIndex        =   23
         Top             =   4140
         Width           =   720
      End
      Begin VB.Label lbl_static_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   465
         Width           =   800
      End
      Begin VB.Label lbl_static_state 
         Height          =   220
         Left            =   1300
         TabIndex        =   32
         Top             =   760
         Width           =   5500
      End
      Begin VB.Label lbl_extern_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   1725
         Width           =   800
      End
      Begin VB.Label lbl_extern_state 
         Height          =   220
         Left            =   1300
         TabIndex        =   12
         Top             =   2020
         Width           =   5500
      End
      Begin VB.Label lbl_tool 
         Caption         =   "Arguments for:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2565
         Width           =   1300
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   4680
      Index           =   1
      Left            =   120
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      Begin VB.OptionButton opt_dll 
         Caption         =   "ActiveX DLL"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   150
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.OptionButton opt_dll 
         Caption         =   "Standard DLL"
         Height          =   255
         Index           =   1
         Left            =   2300
         TabIndex        =   40
         Top             =   150
         Width           =   2000
      End
      Begin VB.TextBox txt_filter 
         Height          =   315
         Left            =   1700
         TabIndex        =   42
         Top             =   570
         Width           =   1720
      End
      Begin VB.ListBox lst_funcs 
         Height          =   2400
         Left            =   120
         MultiSelect     =   2  'Erweitert
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   960
         Width           =   3300
      End
      Begin VB.ListBox lst_exports 
         Height          =   2400
         Left            =   4200
         MultiSelect     =   2  'Erweitert
         TabIndex        =   45
         Top             =   960
         Width           =   3500
      End
      Begin VB.CommandButton btn_exp_add 
         Caption         =   ">"
         Height          =   345
         Left            =   3500
         TabIndex        =   46
         Top             =   1100
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_addall 
         Caption         =   ">>"
         Height          =   345
         Left            =   3500
         TabIndex        =   47
         Top             =   1550
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_remove 
         Caption         =   "X"
         Height          =   345
         Left            =   3500
         TabIndex        =   48
         Top             =   2000
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_clear 
         Caption         =   "Clear"
         Height          =   345
         Left            =   3500
         TabIndex        =   49
         Top             =   2450
         Width           =   600
      End
      Begin VB.CommandButton btn_rescan 
         Caption         =   "Rescan OBJs"
         Height          =   345
         Left            =   120
         TabIndex        =   50
         Top             =   3460
         Width           =   1300
      End
      Begin VB.CommandButton btn_exp_manual 
         Caption         =   "Add name..."
         Height          =   345
         Left            =   4200
         TabIndex        =   51
         Top             =   3460
         Width           =   1300
      End
      Begin VB.TextBox txt_def 
         Height          =   315
         Left            =   1050
         TabIndex        =   53
         Top             =   3860
         Width           =   4700
      End
      Begin VB.CommandButton btn_def_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   5800
         TabIndex        =   54
         Top             =   3860
         Width           =   375
      End
      Begin VB.CommandButton btn_createdef 
         Caption         =   "Create .DEF"
         Height          =   345
         Left            =   6300
         TabIndex        =   55
         Top             =   3845
         Width           =   1400
      End
      Begin VB.CheckBox check_autodef 
         Caption         =   "Regenerate the .DEF automatically on compile (Standard DLL)"
         Height          =   240
         Left            =   1050
         TabIndex        =   56
         Top             =   4260
         Value           =   1  'Aktiviert
         Width           =   6000
      End
      Begin VB.Label lbl_funcs 
         Caption         =   "Functions (filter):"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   610
         Width           =   1500
      End
      Begin VB.Label lbl_exports 
         Caption         =   "Exports (.DEF)"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   610
         Width           =   3000
      End
      Begin VB.Label lbl_def 
         Caption         =   ".DEF file:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   3905
         Width           =   900
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   4680
      Index           =   2
      Left            =   120
      TabIndex        =   57
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      Begin VB.CheckBox check_icon 
         Caption         =   "Replace the VB6 icon (VBIcon)"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   7400
      End
      Begin VB.TextBox txt_icon_obj 
         Height          =   315
         Left            =   1300
         TabIndex        =   60
         Top             =   420
         Width           =   4300
      End
      Begin VB.CommandButton btn_icon_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5700
         TabIndex        =   61
         Top             =   420
         Width           =   600
      End
      Begin VB.CommandButton btn_icon_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   6350
         TabIndex        =   62
         Top             =   420
         Width           =   450
      End
      Begin VB.TextBox txt_icon_file 
         Height          =   315
         Left            =   1300
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1080
         Width           =   4300
      End
      Begin VB.CommandButton btn_icon_pick 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   5700
         TabIndex        =   66
         Top             =   1080
         Width           =   1100
      End
      Begin VB.CommandButton btn_icon_clear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   6850
         TabIndex        =   67
         Top             =   1080
         Width           =   750
      End
      Begin VB.PictureBox pic_icon 
         AutoRedraw      =   -1  'True
         Height          =   2430
         Left            =   480
         ScaleHeight     =   158
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   158
         TabIndex        =   68
         Top             =   1500
         Width           =   2430
      End
      Begin VB.CheckBox check_res 
         Caption         =   "256 x 256"
         Height          =   255
         Index           =   0
         Left            =   3200
         TabIndex        =   71
         Top             =   1800
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "128 x 128"
         Height          =   255
         Index           =   1
         Left            =   3200
         TabIndex        =   72
         Top             =   2100
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "64 x 64"
         Height          =   255
         Index           =   2
         Left            =   3200
         TabIndex        =   73
         Top             =   2400
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "48 x 48"
         Height          =   255
         Index           =   3
         Left            =   4700
         TabIndex        =   74
         Top             =   1800
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "32 x 32"
         Height          =   255
         Index           =   4
         Left            =   4700
         TabIndex        =   75
         Top             =   2100
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "16 x 16"
         Height          =   255
         Index           =   5
         Left            =   4700
         TabIndex        =   76
         Top             =   2400
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "1 bit"
         Height          =   255
         Index           =   0
         Left            =   3200
         TabIndex        =   78
         Top             =   3150
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "4 bit"
         Height          =   255
         Index           =   1
         Left            =   3200
         TabIndex        =   79
         Top             =   3450
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "8 bit"
         Height          =   255
         Index           =   2
         Left            =   3200
         TabIndex        =   80
         Top             =   3750
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "24 bit"
         Height          =   255
         Index           =   3
         Left            =   4700
         TabIndex        =   81
         Top             =   3150
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "32 bit"
         Height          =   255
         Index           =   4
         Left            =   4700
         TabIndex        =   24
         Top             =   3450
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_stretch 
         Caption         =   "Stretch (ignore the aspect ratio)"
         Height          =   255
         Left            =   3200
         TabIndex        =   25
         Top             =   4100
         Width           =   4400
      End
      Begin VB.CheckBox check_remove 
         Caption         =   "Only remove the VB6 icon (add no new icon)"
         Height          =   255
         Left            =   3200
         TabIndex        =   26
         Top             =   4350
         Width           =   4400
      End
      Begin VB.Label lbl_icon_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   465
         Width           =   800
      End
      Begin VB.Label lbl_icon_state 
         Height          =   220
         Left            =   1300
         TabIndex        =   63
         Top             =   760
         Width           =   5500
      End
      Begin VB.Label lbl_icon_file 
         Caption         =   "Icon file:"
         Height          =   255
         Left            =   480
         TabIndex        =   64
         Top             =   1125
         Width           =   800
      End
      Begin VB.Label lbl_icon_size 
         Caption         =   "No icon file selected"
         Height          =   255
         Left            =   480
         TabIndex        =   69
         Top             =   3990
         Width           =   2430
      End
      Begin VB.Label lbl_res 
         Caption         =   "Resolutions (up to the original size):"
         Height          =   255
         Left            =   3200
         TabIndex        =   70
         Top             =   1500
         Width           =   4300
      End
      Begin VB.Label lbl_depth 
         Caption         =   "Color depth:"
         Height          =   255
         Left            =   3200
         TabIndex        =   77
         Top             =   2850
         Width           =   4300
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   4680
      Index           =   3
      Left            =   120
      TabIndex        =   82
      Top             =   960
      Visible         =   0   'False
      Width           =   7800
      Begin VB.TextBox txt_info 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3900
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   83
         Top             =   120
         Width           =   7560
      End
      Begin VB.CommandButton btn_info_refresh 
         Caption         =   "Refresh"
         Height          =   345
         Left            =   120
         TabIndex        =   84
         Top             =   4140
         Width           =   1200
      End
      Begin VB.CommandButton btn_info_copy 
         Caption         =   "Copy"
         Height          =   345
         Left            =   1400
         TabIndex        =   85
         Top             =   4140
         Width           =   1200
      End
   End
   Begin VB.CheckBox check_cmd 
      Caption         =   "Run in CMD"
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   5730
      Width           =   1500
   End
   Begin VB.CommandButton btn_preview 
      Caption         =   "Copy commands"
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   6360
      Width           =   1500
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "Compile"
      Height          =   375
      Left            =   5640
      TabIndex        =   89
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton btn_exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   90
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lbl_cfg 
      Caption         =   "Config:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   735
   End
   Begin VB.Label lbl_status 
      Height          =   255
      Left            =   120
      TabIndex        =   87
      Top             =   6030
      Width           =   7800
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
      Begin VB.Menu mnu_tool_icon 
         Caption         =   "VBIcon.exe"
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
' Build pipeline - no temporary files. main.obj is patched in place by the tools
' (they work completely in memory); an untouched copy of the OBJ is kept in RAM
' and written back if any step fails, so Compile can simply be pressed again.
'
'   1. VBExtern  <main.obj>                              (Declare ... Lib "<extern>")
'   2. VBIcon    <main.obj> <icon> /RES: /DEPTH:         (replace the VB6 icon)
'   3. VBStatic  [/DLL:x] <msvbvm60.dll> <main.obj>      (embed the runtime)
'   4. LINK2.EXE
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

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, ByRef w As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, ByRef h As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Long

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
Private Const TOOL_ICON As Long = 3

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
Private mToolIcon As String
Private mRuntime As String

Private mFnName() As String                'functions found in the OBJs
Private mFnScope() As String
Private mFnDecor() As String
Private mFnCount As Long

Private mBkPath() As String                'in-memory copies of the OBJs (restored when a build fails)
Private mBkData As Collection
Private mBkCount As Long

Private mGdipToken As Long                 'GDI+ (icon preview)
Private mSrcW As Long                      'size of the icon source image (0 = none, -1 = unknown)
Private mSrcH As Long

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
    Set mBkData = New Collection

    cbo_tool.AddItem "link"
    cbo_tool.AddItem "vbstatic"
    cbo_tool.AddItem "vbextern"
    cbo_tool.AddItem "vbicon"

    LoadTools
    StartGdip
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
    If mGdipToken <> 0 Then
        GdiplusShutdown mGdipToken
        mGdipToken = 0
    End If
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

    Dim i As Long

    For i = 0 To 3
        fra_tab(i).Visible = (i = Index)
    Next i
    If Index = 3 Then RefreshInfo

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
    Dim ic As Boolean
    Dim rm As Boolean
    Dim i As Long

    s = (check_static.value = vbChecked)
    e = (check_extern.value = vbChecked)
    ic = (check_icon.value = vbChecked)
    rm = (check_remove.value = vbChecked)

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

    UpdateArgsEnabled

    'icon tab
    lbl_icon_obj.Enabled = ic
    txt_icon_obj.Enabled = ic
    btn_icon_auto.Enabled = ic
    btn_icon_browse.Enabled = ic
    lbl_icon_state.Enabled = ic
    check_remove.Enabled = ic
    lbl_icon_file.Enabled = ic And Not rm
    txt_icon_file.Enabled = ic And Not rm
    btn_icon_pick.Enabled = ic And Not rm
    btn_icon_clear.Enabled = ic And Not rm
    check_stretch.Enabled = ic And Not rm
    lbl_res.Enabled = ic And Not rm
    lbl_depth.Enabled = ic And Not rm
    For i = 0 To 4
        check_depth(i).Enabled = ic And Not rm
    Next i
    UpdateResEnable False

    'DLL tab: only for /DLL builds, exports only for standard DLLs
    isDll = IsDllBuild()
    opt_tab(1).Enabled = isDll
    If Not isDll Then
        If opt_tab(1).value Then opt_tab(0).value = True
    End If

    std = isDll And opt_dll(1).value
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
        Case TOOL_STATIC: en = (check_static.value = vbChecked)
        Case TOOL_EXTERN: en = (check_extern.value = vbChecked)
        Case TOOL_ICON: en = (check_icon.value = vbChecked)
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

Private Sub check_cmd_Click()
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
    For i = 0 To 3
        args_list(i).Visible = (i = cbo_tool.ListIndex)
    Next i
    UpdateArgsEnabled

End Sub

Private Sub AddArg(ByVal idx As Long, ByVal Text As String, ByVal checked As Boolean, ByVal flag As Long)

    Dim old As Boolean

    old = mLoading
    mLoading = True
    args_list(idx).AddItem Text
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

Private Sub AddHidden(ByVal Text As String)

    Dim i As Long

    For i = 0 To mHiddenCount - 1
        If StrComp(mHidden(i), Text, vbTextCompare) = 0 Then Exit Sub
    Next i
    ReDim Preserve mHidden(mHiddenCount)
    mHidden(mHiddenCount) = Text
    mHiddenCount = mHiddenCount + 1

End Sub

Private Sub RemoveCmdArg(ByVal Text As String)

    Dim i As Long

    For i = args_list(TOOL_LINK).ListCount - 1 To 0 Step -1
        If args_list(TOOL_LINK).ItemData(i) = FLAG_CMD Then
            If StrComp(args_list(TOOL_LINK).List(i), Text, vbTextCompare) = 0 Then
                args_list(TOOL_LINK).RemoveItem i
            End If
        End If
    Next i

End Sub

Private Sub SetCmdArgOff(ByVal Text As String)

    Dim i As Long

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).ItemData(i) = FLAG_CMD Then
            If StrComp(args_list(TOOL_LINK).List(i), Text, vbTextCompare) = 0 Then
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
Private Function IsObjToken(ByVal token As String) As Boolean

    Dim p As String

    p = StripQ(token)
    If Len(p) < 5 Then Exit Function
    If Left$(p, 1) = "/" Or Left$(p, 1) = "-" Then Exit Function
    IsObjToken = (LCase$(Right$(p, 4)) = ".obj")

End Function

Private Function AddUnique(c As Collection, ByVal Text As String) As Boolean
    On Error Resume Next
    c.Add Text, LCase$(Text)
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

    Dim m As String
    Dim st As String

    m = FindMainObj()
    If Len(m) > 0 Then
        st = "auto-detected (defines ___vbaS)"
    Else
        st = "auto detection failed"
    End If

    txt_static_obj.Text = m
    txt_extern_obj.Text = m
    txt_icon_obj.Text = m
    lbl_static_state.Caption = st
    lbl_extern_state.Caption = st
    lbl_icon_state.Caption = st

End Sub

'The main OBJ is the one that defines ___vbaS. No guessing: empty when none does.
Private Function FindMainObj() As String

    Dim objs As Collection
    Dim v As Variant

    Set objs = ObjsFromList()
    For Each v In objs
        If IsMainObj(CStr(v)) Then
            FindMainObj = CStr(v)
            Exit Function
        End If
    Next v

End Function

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
    Dim d As String

    For i = 0 To mCmdCount - 1
        If IsObjToken(mCmdArgs(i)) Then
            If IsMainObj(StripQ(mCmdArgs(i))) Then
                m = StripQ(mCmdArgs(i))
                Exit For
            End If
        End If
    Next i

    If Len(m) > 0 Then
        d = DirOf(m)
        If Len(d) = 0 Then d = CurDir$
        mCfgDir = d
    Else
        mCfgDir = App.path                       'main.obj not found: config next to this exe
    End If
    mCfgFile = AddBS(mCfgDir) & "vblinker.ini"

End Sub

Private Function ListKey(ByVal k As Long) As String
    Select Case k
        Case TOOL_LINK: ListKey = "Link"
        Case TOOL_STATIC: ListKey = "Static"
        Case TOOL_EXTERN: ListKey = "Extern"
        Case Else: ListKey = "IconArgs"
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
    Dim lst As String

    old = mLoading
    mLoading = True
    sec = "Config." & cfgName
    f = mCfgFile

    'start from the command line of VB6
    For k = 0 To 3
        args_list(k).Clear
    Next k
    mHiddenCount = 0
    For i = 0 To mCmdCount - 1
        AddArg TOOL_LINK, mCmdArgs(i), True, FLAG_CMD
    Next i

    'defaults
    check_static.value = vbUnchecked
    check_extern.value = vbUnchecked
    check_icon.value = vbUnchecked
    check_remove.value = vbUnchecked
    check_stretch.value = vbUnchecked
    check_autofix.value = vbChecked
    check_cmd.value = vbUnchecked
    check_autodef.value = vbChecked
    For i = 0 To 5
        check_res(i).value = vbChecked
    Next i
    For i = 0 To 4
        check_depth(i).value = IIf(i = 4, vbChecked, vbUnchecked)
    Next i
    opt_dll(0).value = True
    lst_exports.Clear
    txt_def.Text = ""
    txt_icon_file.Text = ""

    If IniHas(f, sec) Then

        check_static.value = BoolCheck(IniLong(f, sec, "Static", 0))
        check_extern.value = BoolCheck(IniLong(f, sec, "Extern", 0))
        check_icon.value = BoolCheck(IniLong(f, sec, "Icon", 0))
        check_remove.value = BoolCheck(IniLong(f, sec, "IconRemove", 0))
        check_stretch.value = BoolCheck(IniLong(f, sec, "IconStretch", 0))
        check_autofix.value = BoolCheck(IniLong(f, sec, "AutoFix", 1))
        check_cmd.value = BoolCheck(IniLong(f, sec, "RunInCmd", 0))
        check_autodef.value = BoolCheck(IniLong(f, sec, "AutoDef", 1))
        If IniLong(f, sec, "DllMode", 0) = 1 Then opt_dll(1).value = True Else opt_dll(0).value = True
        txt_def.Text = IniGet(f, sec, "DefFile", "")
        txt_icon_file.Text = Dec(IniGet(f, sec, "IconFile", ""))

        lst = IniGet(f, sec, "IconRes", "*")
        If lst <> "*" Then
            For i = 0 To 5
                check_res(i).value = BoolCheck(IIf(InList(lst, ResSize(i)), 1, 0))
            Next i
        End If
        lst = IniGet(f, sec, "IconDepth", "*")
        If lst <> "*" Then
            For i = 0 To 4
                check_depth(i).value = BoolCheck(IIf(InList(lst, DepthBits(i)), 1, 0))
            Next i
        End If

        n = IniLong(f, sec, "Hide.Count", 0)
        For i = 0 To n - 1
            txt = Dec(IniGet(f, sec, "Hide." & i, ""))
            If Len(txt) > 0 Then
                AddHidden txt
                RemoveCmdArg txt
            End If
        Next i

        For k = 0 To 3
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
    SetIconFile Trim$(txt_icon_file.Text), False

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
    If WritePrivateProfileString(sec, "Version", "2", f) = 0 Then
        MsgBox "Cannot write the config file:" & vbCrLf & f, vbExclamation
        Exit Function
    End If

    IniPut f, sec, "Static", IIf(check_static.value = vbChecked, "1", "0")
    IniPut f, sec, "Extern", IIf(check_extern.value = vbChecked, "1", "0")
    IniPut f, sec, "Icon", IIf(check_icon.value = vbChecked, "1", "0")
    IniPut f, sec, "IconRemove", IIf(check_remove.value = vbChecked, "1", "0")
    IniPut f, sec, "IconStretch", IIf(check_stretch.value = vbChecked, "1", "0")
    IniPut f, sec, "IconFile", Enc(Trim$(txt_icon_file.Text))
    IniPut f, sec, "IconRes", SavedResList()
    IniPut f, sec, "IconDepth", SavedDepthList()
    IniPut f, sec, "AutoFix", IIf(check_autofix.value = vbChecked, "1", "0")
    IniPut f, sec, "RunInCmd", IIf(check_cmd.value = vbChecked, "1", "0")
    IniPut f, sec, "AutoDef", IIf(check_autodef.value = vbChecked, "1", "0")
    IniPut f, sec, "DllMode", IIf(opt_dll(1).value, "1", "0")
    IniPut f, sec, "DefFile", Trim$(txt_def.Text)

    'argument lists: user entries in full, command line entries only when unchecked
    For k = 0 To 3
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
    AppIni = AddBS(App.path) & "vblinker.ini"
End Function

Private Function DefaultRuntime() As String

    Dim p As String

    p = Environ$("windir") & "\SysWOW64\msvbvm60.dll"
    If Not FileExists(p) Then p = Environ$("windir") & "\System32\msvbvm60.dll"
    DefaultRuntime = p

End Function

Private Sub LoadTools()
    mToolLink = IniGet(AppIni(), "Tools", "Link2", AddBS(App.path) & "LINK2.EXE")
    mToolStatic = IniGet(AppIni(), "Tools", "VBStatic", AddBS(App.path) & "VBStatic.exe")
    mToolExtern = IniGet(AppIni(), "Tools", "VBExtern", AddBS(App.path) & "VBExtern.exe")
    mToolIcon = IniGet(AppIni(), "Tools", "VBIcon", AddBS(App.path) & "VBIcon.exe")
    mRuntime = IniGet(AppIni(), "Tools", "Runtime", DefaultRuntime())
End Sub

Private Sub SaveTools()
    IniPut AppIni(), "Tools", "Link2", mToolLink
    IniPut AppIni(), "Tools", "VBStatic", mToolStatic
    IniPut AppIni(), "Tools", "VBExtern", mToolExtern
    IniPut AppIni(), "Tools", "VBIcon", mToolIcon
    IniPut AppIni(), "Tools", "Runtime", mRuntime
End Sub

Private Sub btn_tools_Click()
    mnu_tool_link.Caption = "LINK2.EXE: " & Replace(mToolLink, "&", "&&")
    mnu_tool_static.Caption = "VBStatic.exe: " & Replace(mToolStatic, "&", "&&")
    mnu_tool_extern.Caption = "VBExtern.exe: " & Replace(mToolExtern, "&", "&&")
    mnu_tool_icon.Caption = "VBIcon.exe: " & Replace(mToolIcon, "&", "&&")
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

Private Sub mnu_tool_icon_Click()
    PickTool mToolIcon, "Select VBIcon.exe", "Executables (*.exe)|*.exe|All files (*.*)|*.*"
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

    If IsDllBuild() And opt_dll(1).value And check_autodef.value = vbChecked And lst_exports.ListCount > 0 Then
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
    Dim ic As Boolean
    Dim isDll As Boolean

    s = (check_static.value = vbChecked)
    e = (check_extern.value = vbChecked)
    ic = (check_icon.value = vbChecked)
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
        If Not FileExists(Trim$(txt_extern_obj.Text)) Then
            MsgBox "VBExtern: main.obj was not found." & vbCrLf & "Use Auto or select it manually.", vbExclamation
            Exit Function
        End If
    End If

    If ic Then
        If Not FileExists(mToolIcon) Then
            MsgBox "VBIcon.exe not found:" & vbCrLf & mToolIcon & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If Not FileExists(Trim$(txt_icon_obj.Text)) Then
            MsgBox "VBIcon: main.obj was not found." & vbCrLf & "Use Auto or select it manually.", vbExclamation
            Exit Function
        End If
        If check_remove.value <> vbChecked Then
            If Not FileExists(Trim$(txt_icon_file.Text)) Then
                MsgBox "VBIcon: please select an icon source image (Icon tab).", vbExclamation
                Exit Function
            End If
            If Len(ResList()) = 0 Then
                MsgBox "VBIcon: select at least one resolution.", vbExclamation
                Exit Function
            End If
            If Len(DepthList()) = 0 Then
                MsgBox "VBIcon: select at least one color depth.", vbExclamation
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
            MsgBox "VBStatic: main.obj was not found." & vbCrLf & "Use Auto or select it manually.", vbExclamation
            Exit Function
        End If
    End If

    If isDll And opt_dll(1).value Then
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
'Runs (or, with dryRun, only lists) VBExtern -> VBIcon -> VBStatic -> LINK2. Returns 0 on success.
'The tools patch main.obj in place; the original bytes are kept in memory and restored on failure.
Private Function RunPipeline(ByVal dryRun As Boolean, ByRef report As String) As Long

    Dim isDll As Boolean
    Dim useStatic As Boolean
    Dim useExtern As Boolean
    Dim useIcon As Boolean
    Dim src As String
    Dim cmd As String
    Dim rc As Long
    Dim result As Long

    isDll = IsDllBuild()
    useStatic = (check_static.value = vbChecked)
    useExtern = (check_extern.value = vbChecked)
    useIcon = (check_icon.value = vbChecked)
    report = ""
    ClearBackups

    If Not dryRun Then
        If useExtern Then
            If Not BackupObj(Trim$(txt_extern_obj.Text)) Then result = 1: GoTo Done
        End If
        If useIcon Then
            If Not BackupObj(Trim$(txt_icon_obj.Text)) Then result = 1: GoTo Done
        End If
        If useStatic Then
            If Not BackupObj(Trim$(txt_static_obj.Text)) Then result = 1: GoTo Done
        End If
    End If

    '---- 1. VBExtern ---------------------------------------------------------
    If useExtern Then
        src = Trim$(txt_extern_obj.Text)
        cmd = Q(mToolExtern) & ArgsFor(TOOL_EXTERN, True) & " " & Q(src) & ArgsFor(TOOL_EXTERN, False)
        If dryRun Then
            report = report & cmd & vbCrLf
        Else
            rc = RunStep("VBExtern - " & FileOf(src), cmd, False)
            If rc <> 0 Then
                StepFailed "VBExtern", rc
                result = rc
                GoTo Done
            End If
        End If
    End If

    '---- 2. VBIcon -----------------------------------------------------------
    If useIcon Then
        src = Trim$(txt_icon_obj.Text)
        cmd = Q(mToolIcon) & ArgsFor(TOOL_ICON, True) & " " & Q(src)
        If check_remove.value = vbChecked Then
            cmd = cmd & " /REMOVE"
        Else
            cmd = cmd & " " & Q(Trim$(txt_icon_file.Text)) & " /RES:" & ResList() & " /DEPTH:" & DepthList()
            If check_stretch.value = vbChecked Then cmd = cmd & " /STRETCH"
        End If
        cmd = cmd & ArgsFor(TOOL_ICON, False)
        If dryRun Then
            report = report & cmd & vbCrLf
        Else
            rc = RunStep("VBIcon - " & FileOf(src), cmd, False)
            If rc <> 0 Then
                StepFailed "VBIcon", rc
                result = rc
                GoTo Done
            End If
        End If
    End If

    '---- 3. VBStatic ---------------------------------------------------------
    If useStatic Then
        src = Trim$(txt_static_obj.Text)
        cmd = Q(mToolStatic) & ArgsFor(TOOL_STATIC, True)
        If isDll Then
            If opt_dll(1).value Then cmd = cmd & " /DLL:STANDARD" Else cmd = cmd & " /DLL:ACTIVEX"
        End If
        cmd = cmd & " " & Q(mRuntime) & " " & Q(src) & ArgsFor(TOOL_STATIC, False)
        If dryRun Then
            report = report & cmd & vbCrLf
        Else
            rc = RunStep("VBStatic - " & FileOf(src), cmd, False)
            'VBStatic always exits with 0 - success means the runtime section (.vbrt) is in the OBJ now
            If rc = 0 Then
                If Not ObjHasSection(src, ".vbrt") Then rc = 1
            End If
            If rc <> 0 Then
                StepFailed "VBStatic", rc
                result = rc
                GoTo Done
            End If
        End If
    End If

    '---- 4. link -------------------------------------------------------------
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
    If Not dryRun Then
        If result <> 0 Then RestoreBackups
        ClearBackups
    End If
    RunPipeline = result

End Function

Private Sub StepFailed(ByVal tool As String, ByVal rc As Long)

    Dim msg As String

    If rc = -1 Then
        msg = "Could not start " & tool & "."
    Else
        msg = tool & " failed (exit code " & rc & ")."
    End If
    If check_cmd.value <> vbChecked Then msg = msg & vbCrLf & "Enable 'Run in CMD' to see the output of the tool."
    msg = msg & vbCrLf & "The OBJ files are restored to their original state."
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
    Dim s As String
    Dim entry As String
    Dim doFix As Boolean
    Dim skip As Boolean
    Dim hasEntry As Boolean
    Dim hasNoDef As Boolean

    doFix = useStatic And (check_autofix.value = vbChecked)
    If isDll Then entry = "_VBStaticDllEntry@12" Else entry = "_VBStaticEntry"

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            t = args_list(TOOL_LINK).List(i)
            u = UCase$(StripQ(t))
            skip = False

            If doFix Then
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

'Runs one step and waits for it. Returns the exit code (-1 = could not start).
'Runs one step and waits for it. Returns the exit code (-1 = could not start).
'No script files: in CMD mode one "cmd /c" command line keeps the window open and returns the exit code.
Private Function RunStep(ByVal title As String, ByVal cmdline As String, ByVal isLast As Boolean) As Long

    Dim qq As String
    Dim inner As String

    If check_cmd.value <> vbChecked Then
        RunStep = RunWait(cmdline, SW_SHOWMINNOACTIVE)
        Exit Function
    End If

    qq = Chr$(34)
    inner = "title " & CmdEscape(title) & " & start " & qq & qq & " /wait /b " & cmdline & _
            " & set rc=!errorlevel! & echo. & echo [exit code !rc!] & "
    If isLast Then
        inner = inner & "pause"
    Else
        inner = inner & "(if not !rc!==0 pause)"
    End If
    inner = inner & " & exit /b !rc!"

    RunStep = RunWait(Environ$("ComSpec") & " /v:on /c " & qq & inner & qq, SW_SHOWNORMAL)

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
    s = Replace(s, "!", "^^!")
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


'==============================================================================
' In-memory backup of the OBJs (no temporary files)
'==============================================================================
Private Sub ClearBackups()
    mBkCount = 0
    Erase mBkPath
    Set mBkData = New Collection
End Sub

Private Function BackupObj(ByVal path As String) As Boolean

    Dim b() As Byte
    Dim h As Long
    Dim i As Long
    Dim n As Long

    For i = 0 To mBkCount - 1
        If StrComp(mBkPath(i), path, vbTextCompare) = 0 Then BackupObj = True: Exit Function
    Next i

    On Error GoTo EH
    n = FileLen(path)
    If n <= 0 Then Err.Raise 5, , "file is empty"
    ReDim b(n - 1)
    h = FreeFile
    Open path For Binary Access Read As #h
    Get #h, , b
    Close #h
    h = 0

    ReDim Preserve mBkPath(mBkCount)
    mBkPath(mBkCount) = path
    mBkData.Add b
    mBkCount = mBkCount + 1
    BackupObj = True
    Exit Function

EH:
    On Error Resume Next
    If h <> 0 Then Close #h
    MsgBox "Cannot read " & path & vbCrLf & Err.Description, vbExclamation

End Function

Private Sub RestoreBackups()

    Dim i As Long
    Dim h As Long
    Dim b() As Byte

    For i = 0 To mBkCount - 1
        b = mBkData(i + 1)
        On Error Resume Next
        Kill mBkPath(i)
        h = FreeFile
        Open mBkPath(i) For Binary Access Write As #h
        Put #h, , b
        Close #h
        On Error GoTo 0
    Next i

End Sub

'==============================================================================
' COFF helpers for the loaded object (sections / counters)
'==============================================================================
Private Function ObjSecName(ByVal i As Long) As String

    Dim k As Long
    Dim hd As Long
    Dim s As String

    hd = 20 + 40& * i
    For k = 0 To 7
        If mOb(hd + k) = 0 Then Exit For
        s = s & Chr$(mOb(hd + k))
    Next k
    ObjSecName = s

End Function

Private Function ObjSecText(ByVal i As Long) As String

    Dim hd As Long
    Dim raw As Long
    Dim sz As Long
    Dim k As Long
    Dim s As String

    hd = 20 + 40& * i
    raw = I32(hd + 20)
    sz = I32(hd + 16)
    If raw <= 0 Or sz <= 0 Then Exit Function
    If raw + sz > UBound(mOb) + 1 Then Exit Function
    s = Space$(sz)
    For k = 0 To sz - 1
        Mid$(s, k + 1, 1) = Chr$(mOb(raw + k))
    Next k
    ObjSecText = s

End Function

Private Function ObjHasSection(ByVal path As String, ByVal secName As String) As Boolean

    Dim i As Long
    Dim n As Long

    If Not ObjLoad(path) Then Exit Function
    n = U16(2)
    For i = 0 To n - 1
        If 20 + 40& * (i + 1) > UBound(mOb) + 1 Then Exit For
        If StrComp(ObjSecName(i), secName, vbTextCompare) = 0 Then
            ObjHasSection = True
            Exit Function
        End If
    Next i

End Function

'Number of VB6 Declare thunks in the loaded object (the 25 byte pattern VBExtern patches)
Private Function CountDeclares() As Long

    Dim p As Long
    Dim n As Long

    n = UBound(mOb) - 24
    For p = 0 To n
        If mOb(p) = &HA1 Then
            If mOb(p + 5) = &HB And mOb(p + 6) = &HC0 And mOb(p + 7) = &H74 And mOb(p + 8) = 2 Then
                If mOb(p + 9) = &HFF And mOb(p + 10) = &HE0 And mOb(p + 11) = &H68 And mOb(p + 16) = &HB8 Then
                    If mOb(p + 21) = &HFF And mOb(p + 22) = &HD0 And mOb(p + 23) = &HFF And mOb(p + 24) = &HE0 Then
                        CountDeclares = CountDeclares + 1
                    End If
                End If
            End If
        End If
    Next p

End Function

Private Function EndsWithAtDigits(ByVal s As String) As Boolean

    Dim p As Long
    Dim i As Long

    p = InStrRev(s, "@")
    If p < 2 Or p = Len(s) Then Exit Function
    For i = p + 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    EndsWithAtDigits = True

End Function

'==============================================================================
' Icon tab
'==============================================================================
Private Function ResSize(ByVal i As Long) As Long
    Select Case i
        Case 0: ResSize = 256
        Case 1: ResSize = 128
        Case 2: ResSize = 64
        Case 3: ResSize = 48
        Case 4: ResSize = 32
        Case Else: ResSize = 16
    End Select
End Function

Private Function DepthBits(ByVal i As Long) As Long
    Select Case i
        Case 0: DepthBits = 1
        Case 1: DepthBits = 4
        Case 2: DepthBits = 8
        Case 3: DepthBits = 24
        Case Else: DepthBits = 32
    End Select
End Function

Private Function InList(ByVal lst As String, ByVal n As Long) As Boolean
    InList = (InStr("," & lst & ",", "," & n & ",") > 0)
End Function

'comma separated list of the active sizes / depths (for /RES: and /DEPTH:)
Private Function ResList() As String

    Dim i As Long
    Dim s As String

    For i = 0 To 5
        If check_res(i).value = vbChecked And check_res(i).Enabled Then s = s & "," & ResSize(i)
    Next i
    If Len(s) > 0 Then s = Mid$(s, 2)
    ResList = s

End Function

Private Function DepthList() As String

    Dim i As Long
    Dim s As String

    For i = 0 To 4
        If check_depth(i).value = vbChecked And check_depth(i).Enabled Then s = s & "," & DepthBits(i)
    Next i
    If Len(s) > 0 Then s = Mid$(s, 2)
    DepthList = s

End Function

'for the config file: every checked box, whether currently enabled or not
Private Function SavedResList() As String

    Dim i As Long
    Dim s As String

    For i = 0 To 5
        If check_res(i).value = vbChecked Then s = s & "," & ResSize(i)
    Next i
    SavedResList = Mid$(s & ",", 2)

End Function

Private Function SavedDepthList() As String

    Dim i As Long
    Dim s As String

    For i = 0 To 4
        If check_depth(i).value = vbChecked Then s = s & "," & DepthBits(i)
    Next i
    SavedDepthList = Mid$(s & ",", 2)

End Function

Private Sub StartGdip()

    Dim gi As GdiplusStartupInput

    gi.GdiplusVersion = 1
    On Error Resume Next
    If GdiplusStartup(mGdipToken, gi, 0&) <> 0 Then mGdipToken = 0
    If Err.Number <> 0 Then mGdipToken = 0
    Err.Clear

End Sub

Private Function GdipSize(ByVal path As String, ByRef w As Long, ByRef h As Long) As Boolean

    Dim img As Long

    w = 0
    h = 0
    If mGdipToken = 0 Then Exit Function
    On Error Resume Next
    If GdipLoadImageFromFile(StrPtr(path), img) <> 0 Then Exit Function
    GdipGetImageWidth img, w
    GdipGetImageHeight img, h
    GdipDisposeImage img
    GdipSize = (w > 0 And h > 0)

End Function

Private Sub DrawChecker()

    Dim x As Long
    Dim y As Long
    Dim c As Long

    pic_icon.Cls
    For y = 0 To pic_icon.ScaleHeight - 1 Step 8
        For x = 0 To pic_icon.ScaleWidth - 1 Step 8
            If (((x \ 8) + (y \ 8)) Mod 2) = 0 Then c = RGB(255, 255, 255) Else c = RGB(204, 204, 204)
            pic_icon.Line (x, y)-(x + 7, y + 7), c, BF
        Next x
    Next y

End Sub

Private Sub DrawPreview(ByVal path As String)

    Dim img As Long
    Dim g As Long
    Dim w As Long
    Dim h As Long
    Dim sc As Double
    Dim dw As Long
    Dim dh As Long
    Dim pw As Long
    Dim ph As Long

    DrawChecker
    If mGdipToken <> 0 And Len(path) > 0 Then
        If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
            GdipGetImageWidth img, w
            GdipGetImageHeight img, h
            If w > 0 And h > 0 Then
                pw = pic_icon.ScaleWidth
                ph = pic_icon.ScaleHeight
                sc = pw / w
                If ph / h < sc Then sc = ph / h
                dw = CLng(w * sc)
                dh = CLng(h * sc)
                If dw < 1 Then dw = 1
                If dh < 1 Then dh = 1
                If GdipCreateFromHDC(pic_icon.hdc, g) = 0 Then
                    GdipSetInterpolationMode g, 7
                    GdipDrawImageRectI g, img, (pw - dw) \ 2, (ph - dh) \ 2, dw, dh
                    GdipDeleteGraphics g
                End If
            End If
            GdipDisposeImage img
        End If
    End If
    pic_icon.Refresh

End Sub

Private Sub SetIconFile(ByVal path As String, ByVal resetChecks As Boolean)

    Dim w As Long
    Dim h As Long
    Dim old As Boolean

    old = mLoading
    mLoading = True
    txt_icon_file.Text = path
    mSrcW = 0
    mSrcH = 0

    If Len(path) = 0 Then
        lbl_icon_size.Caption = "No icon file selected"
        DrawPreview ""
    ElseIf Not FileExists(path) Then
        lbl_icon_size.Caption = "File not found"
        DrawPreview ""
    ElseIf mGdipToken = 0 Then
        mSrcW = -1                                  'size unknown: every resolution may be chosen
        mSrcH = -1
        lbl_icon_size.Caption = "GDI+ not available - no preview"
        DrawPreview ""
    ElseIf GdipSize(path, w, h) Then
        mSrcW = w
        mSrcH = h
        lbl_icon_size.Caption = "Original: " & w & " x " & h & " px"
        DrawPreview path
    Else
        lbl_icon_size.Caption = "Cannot read this image"
        DrawPreview ""
    End If

    UpdateResEnable resetChecks
    mLoading = old

End Sub

'A resolution can be selected when it is not larger than the smaller side of the original (16 always)
Private Sub UpdateResEnable(ByVal resetChecks As Boolean)

    Dim i As Long
    Dim m As Long
    Dim en As Boolean
    Dim master As Boolean
    Dim old As Boolean

    old = mLoading
    mLoading = True
    master = (check_icon.value = vbChecked) And (check_remove.value <> vbChecked)
    m = mSrcW
    If mSrcH < m Then m = mSrcH

    For i = 0 To 5
        If mSrcW < 0 Then
            en = True
        ElseIf mSrcW = 0 Then
            en = False
        Else
            en = (ResSize(i) <= m) Or (ResSize(i) = 16)
        End If
        If mSrcW > 0 Then
            If Not en Then
                check_res(i).value = vbUnchecked
            ElseIf resetChecks Then
                check_res(i).value = vbChecked
            End If
        End If
        check_res(i).Enabled = en And master
    Next i
    mLoading = old

End Sub

Private Sub check_icon_Click()
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub check_remove_Click()
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub check_stretch_Click()
    MarkDirty
End Sub

Private Sub check_res_Click(Index As Integer)
    MarkDirty
End Sub

Private Sub check_depth_Click(Index As Integer)
    MarkDirty
End Sub

Private Sub txt_icon_file_Change()
    MarkDirty
End Sub

Private Sub btn_icon_pick_Click()

    Dim c As Collection

    Set c = FileDialog("Select the icon source image", "Images (*.png;*.jpg;*.jpeg;*.bmp;*.ico)|*.png;*.jpg;*.jpeg;*.bmp;*.ico|All files (*.*)|*.*", False, False, txt_icon_file.Text)
    If c Is Nothing Then Exit Sub
    SetIconFile c(1), True
    MarkDirty

End Sub

Private Sub btn_icon_clear_Click()
    SetIconFile "", False
    MarkDirty
End Sub

Private Sub btn_icon_auto_Click()
    AutoDetectObjs
End Sub

Private Sub btn_icon_browse_Click()

    Dim c As Collection

    Set c = FileDialog("Select main.obj", "Object files (*.obj)|*.obj|All files (*.*)|*.*", False, False, txt_icon_obj.Text)
    If c Is Nothing Then Exit Sub
    txt_icon_obj.Text = c(1)
    lbl_icon_state.Caption = "selected manually"

End Sub

'==============================================================================
' Info tab
'==============================================================================
Private Sub btn_info_refresh_Click()
    RefreshInfo
End Sub

Private Sub btn_info_copy_Click()
    Clipboard.Clear
    Clipboard.SetText txt_info.Text
End Sub

Private Function InfoLine(ByVal label As String, ByVal value As String, Optional ByVal indent As Long = 2) As String
    If Len(value) = 0 Then value = "-"
    InfoLine = Space$(indent) & Left$(label & Space$(18), 18) & ": " & value & vbCrLf
End Function

Private Function ToolState(ByVal path As String) As String
    If FileExists(path) Then ToolState = "found    " & path Else ToolState = "MISSING  " & path
End Function

Private Function OnOff(ByVal b As Boolean) As String
    If b Then OnOff = "on" Else OnOff = "off"
End Function

'value of the first checked link argument that starts with prefix (e.g. "/OUT:")
Private Function ArgValue(ByVal prefix As String) As String

    Dim i As Long
    Dim t As String

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            t = args_list(TOOL_LINK).List(i)
            If UCase$(Left$(t, Len(prefix))) = UCase$(prefix) Then
                ArgValue = StripQ(Mid$(t, Len(prefix) + 1))
                Exit Function
            End If
        End If
    Next i

End Function

Private Sub RefreshInfo()

    Dim s As String
    Dim objs As Collection
    Dim v As Variant
    Dim main As String
    Dim kind As String
    Dim i As Long
    Dim nOn As Long
    Dim isDll As Boolean

    isDll = IsDllBuild()
    If isDll Then
        If opt_dll(1).value Then kind = "DLL (Standard)" Else kind = "DLL (ActiveX)"
    Else
        kind = "EXE"
    End If
    main = FindMainObj()
    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then nOn = nOn + 1
    Next i

    s = "PROJECT" & vbCrLf
    s = s & InfoLine("Config", mCurCfg)
    s = s & InfoLine("Config file", mCfgFile)
    s = s & InfoLine("Build type", kind)
    s = s & InfoLine("Output file", ArgValue("/OUT:"))
    s = s & InfoLine("Entry point", ArgValue("/ENTRY:"))
    s = s & InfoLine("Subsystem", ArgValue("/SUBSYSTEM:"))
    s = s & InfoLine("Image base", ArgValue("/BASE:"))
    If isDll Then
        s = s & InfoLine("DEF file", ArgValue("/DEF:"))
        If opt_dll(1).value Then s = s & InfoLine("Exports (.DEF)", CStr(lst_exports.ListCount))
    End If
    s = s & InfoLine("Link arguments", nOn & " active of " & args_list(TOOL_LINK).ListCount)
    If Len(main) > 0 Then
        s = s & InfoLine("main.obj", main)
        s = s & InfoLine("main.obj size", Format$(FileLen(main), "#,##0") & " bytes")
    Else
        s = s & InfoLine("main.obj", "auto detection failed")
    End If

    s = s & vbCrLf & "BUILD STEPS" & vbCrLf
    s = s & InfoLine("1. VBExtern", OnOff(check_extern.value = vbChecked))
    If check_icon.value = vbChecked Then
        If check_remove.value = vbChecked Then
            s = s & InfoLine("2. VBIcon", "on - remove the VB6 icon only")
        Else
            s = s & InfoLine("2. VBIcon", "on - " & FileOf(Trim$(txt_icon_file.Text)) & "  sizes " & ResList() & "  depth " & DepthList())
        End If
    Else
        s = s & InfoLine("2. VBIcon", "off")
    End If
    If check_static.value = vbChecked Then
        s = s & InfoLine("3. VBStatic", "on - runtime " & mRuntime)
    Else
        s = s & InfoLine("3. VBStatic", "off")
    End If
    s = s & InfoLine("4. LINK2", "always")

    s = s & vbCrLf & "TOOLS" & vbCrLf
    s = s & InfoLine("LINK2.EXE", ToolState(mToolLink))
    s = s & InfoLine("VBExtern.exe", ToolState(mToolExtern))
    s = s & InfoLine("VBIcon.exe", ToolState(mToolIcon))
    s = s & InfoLine("VBStatic.exe", ToolState(mToolStatic))
    s = s & InfoLine("MSVBVM60.DLL", ToolState(mRuntime))

    Set objs = ObjsFromList()
    s = s & vbCrLf & "OBJECT FILES (" & objs.count & ")" & vbCrLf
    For Each v In objs
        s = s & ObjInfoText(CStr(v)) & vbCrLf
    Next v

    txt_info.Text = s
    txt_info.SelStart = 0

End Sub

Private Function ObjInfoText(ByVal path As String) As String

    Dim s As String
    Dim i As Long
    Dim p As Long
    Dim nSec As Long
    Dim secn As Long
    Dim nm As String
    Dim dn As String
    Dim ds As String
    Dim names As String
    Dim exps As String
    Dim undef As String
    Dim txt As String
    Dim tok As Variant
    Dim nFn As Long
    Dim nUndef As Long
    Dim isMain As Boolean
    Dim hasRsrc As Boolean
    Dim patched As Boolean
    Dim usesRt As Boolean

    If Not ObjLoad(path) Then
        ObjInfoText = "  " & FileOf(path) & vbCrLf & "      not readable as an i386 COFF object: " & path & vbCrLf
        Exit Function
    End If

    nSec = U16(2)
    On Error Resume Next
    For i = 0 To nSec - 1
        nm = ObjSecName(i)
        If i < 12 Then names = names & nm & " "
        If nm = ".vbrt" Then patched = True
        If Left$(nm, 5) = ".rsrc" Then hasRsrc = True
        If nm = ".drectve" Then
            txt = ObjSecText(i)
            For Each tok In Split(txt, " ")
                If LCase$(Left$(tok, 8)) = "-export:" Or LCase$(Left$(tok, 8)) = "/export:" Then exps = exps & Mid$(tok, 9) & " "
            Next tok
        End If
    Next i

    i = 0
    Do While i < mObSymCnt
        p = mObSymOff + i * 18&
        If p + 17 > UBound(mOb) Then Exit Do
        If mOb(p + 16) = 2 Then
            secn = I16(p + 12)
            If secn > 0 Then
                nm = ObjSymName(i)
                If nm = "___vbaS" Then isMain = True
                If Demangle(nm, dn, ds) Then nFn = nFn + 1
            ElseIf secn = 0 Then
                nm = ObjSymName(i)
                If Left$(nm, 6) <> "__imp_" Then
                    If EndsWithAtDigits(nm) Then
                        nUndef = nUndef + 1
                        If nUndef <= 30 Then undef = undef & nm & " "
                    End If
                End If
            End If
        End If
        i = i + 1 + mOb(p + 17)
    Loop
    usesRt = (InStr(1, StrConv(mOb, vbUnicode), "msvbvm60", vbTextCompare) > 0)
    On Error GoTo 0

    s = "  " & FileOf(path)
    If isMain Then s = s & "     <== MAIN OBJECT"
    s = s & vbCrLf
    s = s & InfoLine("path", path, 6)
    s = s & InfoLine("size", Format$(UBound(mOb) + 1, "#,##0") & " bytes", 6)
    s = s & InfoLine("sections", nSec & "  " & Trim$(names), 6)
    s = s & InfoLine("symbols", mObSymCnt & "   (public VB functions: " & nFn & ")", 6)
    s = s & InfoLine("Declare thunks", CStr(CountDeclares()), 6)
    If Len(Trim$(undef)) > 0 Then
        s = s & InfoLine("undefined externs", nUndef & "  " & Trim$(undef), 6)
    End If
    If Len(Trim$(exps)) > 0 Then s = s & InfoLine("-export: directives", Trim$(exps), 6)
    s = s & InfoLine("resources", IIf(hasRsrc, "yes (.rsrc)", "none"), 6)
    s = s & InfoLine("msvbvm60 import", IIf(usesRt, "referenced", "not referenced"), 6)
    If isMain Then s = s & InfoLine("VBStatic state", IIf(patched, "runtime already embedded (.vbrt)", "not patched"), 6)
    ObjInfoText = s

End Function

VERSION 5.00
Begin VB.Form frmLinker 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Linker"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   Icon            =   "Linker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
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
      Height          =   7040
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   7800
      Begin VB.TextBox txt_main_obj 
         Height          =   315
         Left            =   1050
         TabIndex        =   34
         Top             =   320
         Width           =   4700
      End
      Begin VB.CommandButton btn_main_auto 
         Caption         =   "Auto"
         Height          =   315
         Left            =   5800
         TabIndex        =   35
         Top             =   320
         Width           =   600
      End
      Begin VB.CommandButton btn_main_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   6450
         TabIndex        =   36
         Top             =   320
         Width           =   450
      End
      Begin VB.CheckBox check_static 
         Caption         =   "Embed MSVBVM60.DLL into build (VBStatic)"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1140
         Width           =   7400
      End
      Begin VB.CheckBox check_autofix 
         Caption         =   "Adjust link args for VBStatic (entry point, /NODEFAULTLIB)"
         Height          =   240
         Left            =   480
         TabIndex        =   60
         Top             =   1460
         Value           =   1  'Aktiviert
         Width           =   7200
      End
      Begin VB.CheckBox check_noicon 
         Caption         =   "Create empty icon ressource (Use when removing icon)"
         Height          =   240
         Left            =   480
         TabIndex        =   61
         Top             =   1800
         Width           =   7200
      End
      Begin VB.CheckBox check_extern 
         Caption         =   "Support for external symbols (VBExtern)"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   2200
         Width           =   7400
      End
      Begin VB.CheckBox check_manifest 
         Caption         =   "Create Manifest (VBMF)"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   2600
         Width           =   7400
      End
      Begin VB.CheckBox check_vstyles 
         Caption         =   "Enable Visual Styles"
         Height          =   255
         Left            =   480
         TabIndex        =   92
         Top             =   2980
         Width           =   2300
      End
      Begin VB.CheckBox check_dpi 
         Caption         =   "DPI Aware"
         Height          =   255
         Left            =   2900
         TabIndex        =   11
         Top             =   2980
         Width           =   1500
      End
      Begin VB.CheckBox check_segheap 
         Caption         =   "Use Segment Heap"
         Height          =   255
         Left            =   4500
         TabIndex        =   12
         Top             =   2980
         Width           =   2400
      End
      Begin VB.ComboBox cbo_uac 
         Height          =   315
         Left            =   1100
         Style           =   2  'Dropdown-Liste
         TabIndex        =   14
         Top             =   3360
         Width           =   2300
      End
      Begin VB.ComboBox cbo_os 
         Height          =   315
         Left            =   4900
         Style           =   2  'Dropdown-Liste
         TabIndex        =   16
         Top             =   3360
         Width           =   2300
      End
      Begin VB.TextBox txt_manifest 
         Height          =   315
         Left            =   1500
         TabIndex        =   18
         Top             =   3860
         Width           =   3900
      End
      Begin VB.CommandButton btn_manifest_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   5450
         TabIndex        =   19
         Top             =   3860
         Width           =   375
      End
      Begin VB.CommandButton btn_createmanifest 
         Caption         =   "Create manifest"
         Height          =   345
         Left            =   5900
         TabIndex        =   20
         Top             =   3845
         Width           =   1800
      End
      Begin VB.CheckBox check_automanifest 
         Caption         =   "Regenerate the manifest automatically on compile"
         Height          =   240
         Left            =   1500
         TabIndex        =   21
         Top             =   4330
         Value           =   1  'Aktiviert
         Width           =   6200
      End
      Begin VB.ComboBox cbo_tool 
         Height          =   315
         Left            =   1450
         Style           =   2  'Dropdown-Liste
         TabIndex        =   23
         Top             =   4740
         Width           =   2400
      End
      Begin VB.ListBox args_list 
         Height          =   1635
         Index           =   0
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   24
         Top             =   5240
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1635
         Index           =   1
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   25
         Top             =   5240
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1635
         Index           =   2
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   26
         Top             =   5240
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1635
         Index           =   3
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   27
         Top             =   5240
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.ListBox args_list 
         Height          =   1635
         Index           =   4
         Left            =   120
         Style           =   1  'Kontrollkästchen
         TabIndex        =   28
         Top             =   5240
         Visible         =   0   'False
         Width           =   6800
      End
      Begin VB.CommandButton btn_add 
         Caption         =   "+"
         Height          =   285
         Left            =   7000
         TabIndex        =   29
         Top             =   5240
         Width           =   720
      End
      Begin VB.CommandButton btn_remove 
         Caption         =   "-"
         Height          =   285
         Left            =   7000
         TabIndex        =   30
         Top             =   5580
         Width           =   720
      End
      Begin VB.CommandButton btn_addfile 
         Caption         =   "File..."
         Height          =   285
         Left            =   7000
         TabIndex        =   31
         Top             =   5920
         Width           =   720
      End
      Begin VB.CommandButton btn_up 
         Caption         =   "Up"
         Height          =   285
         Left            =   7000
         TabIndex        =   32
         Top             =   6260
         Width           =   720
      End
      Begin VB.CommandButton btn_down 
         Caption         =   "Down"
         Height          =   285
         Left            =   7000
         TabIndex        =   33
         Top             =   6600
         Width           =   720
      End
      Begin VB.Label lbl_main_obj 
         Caption         =   "main.obj:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   365
         Width           =   900
      End
      Begin VB.Label lbl_main_state 
         Height          =   220
         Left            =   1050
         TabIndex        =   37
         Top             =   740
         Width           =   5800
      End
      Begin VB.Label lbl_uac 
         Caption         =   "UAC:"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   3405
         Width           =   600
      End
      Begin VB.Label lbl_os 
         Caption         =   "Supported OS:"
         Height          =   255
         Left            =   3700
         TabIndex        =   15
         Top             =   3405
         Width           =   1200
      End
      Begin VB.Label lbl_manifest 
         Caption         =   "Manifest file:"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   3905
         Width           =   1000
      End
      Begin VB.Label lbl_tool 
         Caption         =   "Arguments for:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   4785
         Width           =   1300
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   7040
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
         Top             =   350
         Value           =   -1  'True
         Width           =   2000
      End
      Begin VB.OptionButton opt_dll 
         Caption         =   "Standard DLL"
         Height          =   255
         Index           =   1
         Left            =   2300
         TabIndex        =   40
         Top             =   350
         Width           =   2000
      End
      Begin VB.TextBox txt_filter 
         Height          =   315
         Left            =   1700
         TabIndex        =   42
         Top             =   850
         Width           =   1720
      End
      Begin VB.ListBox lst_funcs 
         Height          =   2985
         Left            =   120
         MultiSelect     =   2  'Erweitert
         Sorted          =   -1  'True
         TabIndex        =   44
         Top             =   1320
         Width           =   3300
      End
      Begin VB.ListBox lst_exports 
         Height          =   2985
         Left            =   4200
         MultiSelect     =   2  'Erweitert
         TabIndex        =   45
         Top             =   1320
         Width           =   3500
      End
      Begin VB.CommandButton btn_exp_add 
         Caption         =   ">"
         Height          =   345
         Left            =   3500
         TabIndex        =   46
         Top             =   1460
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_addall 
         Caption         =   ">>"
         Height          =   345
         Left            =   3500
         TabIndex        =   47
         Top             =   1910
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_remove 
         Caption         =   "X"
         Height          =   345
         Left            =   3500
         TabIndex        =   48
         Top             =   2360
         Width           =   600
      End
      Begin VB.CommandButton btn_exp_clear 
         Caption         =   "Clear"
         Height          =   345
         Left            =   3500
         TabIndex        =   49
         Top             =   2810
         Width           =   600
      End
      Begin VB.CommandButton btn_rescan 
         Caption         =   "Rescan OBJs"
         Height          =   345
         Left            =   120
         TabIndex        =   50
         Top             =   4500
         Width           =   1300
      End
      Begin VB.CommandButton btn_exp_manual 
         Caption         =   "Add name..."
         Height          =   345
         Left            =   4200
         TabIndex        =   51
         Top             =   4500
         Width           =   1300
      End
      Begin VB.TextBox txt_def 
         Height          =   315
         Left            =   1050
         TabIndex        =   53
         Top             =   4980
         Width           =   4700
      End
      Begin VB.CommandButton btn_def_browse 
         Caption         =   "..."
         Height          =   315
         Left            =   5800
         TabIndex        =   54
         Top             =   4980
         Width           =   375
      End
      Begin VB.CommandButton btn_createdef 
         Caption         =   "Create .DEF"
         Height          =   345
         Left            =   6300
         TabIndex        =   55
         Top             =   4965
         Width           =   1400
      End
      Begin VB.CheckBox check_autodef 
         Caption         =   "Regenerate the .DEF automatically on compile (Standard DLL)"
         Height          =   240
         Left            =   1050
         TabIndex        =   56
         Top             =   5460
         Value           =   1  'Aktiviert
         Width           =   6000
      End
      Begin VB.Label lbl_funcs 
         Caption         =   "Functions (filter):"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   890
         Width           =   1500
      End
      Begin VB.Label lbl_exports 
         Caption         =   "Exports (.DEF)"
         Height          =   255
         Left            =   4200
         TabIndex        =   43
         Top             =   890
         Width           =   3000
      End
      Begin VB.Label lbl_def 
         Caption         =   ".DEF file:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   5025
         Width           =   900
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   7040
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
         Top             =   320
         Width           =   7400
      End
      Begin VB.CheckBox check_remove 
         Caption         =   "Remove icon"
         Height          =   255
         Left            =   480
         TabIndex        =   93
         Top             =   700
         Width           =   2600
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
         Top             =   1580
         Width           =   2430
      End
      Begin VB.CheckBox check_res 
         Caption         =   "256 x 256"
         Height          =   255
         Index           =   0
         Left            =   3200
         TabIndex        =   71
         Top             =   1960
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "128 x 128"
         Height          =   255
         Index           =   1
         Left            =   3200
         TabIndex        =   72
         Top             =   2320
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "64 x 64"
         Height          =   255
         Index           =   2
         Left            =   3200
         TabIndex        =   73
         Top             =   2680
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "48 x 48"
         Height          =   255
         Index           =   3
         Left            =   4700
         TabIndex        =   74
         Top             =   1960
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "32 x 32"
         Height          =   255
         Index           =   4
         Left            =   4700
         TabIndex        =   75
         Top             =   2320
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_res 
         Caption         =   "16 x 16"
         Height          =   255
         Index           =   5
         Left            =   4700
         TabIndex        =   76
         Top             =   2680
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "1 bit"
         Height          =   255
         Index           =   0
         Left            =   3200
         TabIndex        =   78
         Top             =   3560
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "4 bit"
         Height          =   255
         Index           =   1
         Left            =   3200
         TabIndex        =   79
         Top             =   3920
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "8 bit"
         Height          =   255
         Index           =   2
         Left            =   3200
         TabIndex        =   80
         Top             =   4280
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "24 bit"
         Height          =   255
         Index           =   3
         Left            =   4700
         TabIndex        =   81
         Top             =   3560
         Width           =   1400
      End
      Begin VB.CheckBox check_depth 
         Caption         =   "32 bit"
         Height          =   255
         Index           =   4
         Left            =   4700
         TabIndex        =   94
         Top             =   3920
         Value           =   1  'Aktiviert
         Width           =   1400
      End
      Begin VB.CheckBox check_stretch 
         Caption         =   "Stretch (ignore the aspect ratio)"
         Height          =   255
         Left            =   3200
         TabIndex        =   95
         Top             =   4695
         Width           =   4400
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
         Height          =   600
         Left            =   480
         TabIndex        =   69
         Top             =   4100
         Width           =   2600
      End
      Begin VB.Label lbl_res 
         Caption         =   "Resolutions (up to the original size):"
         Height          =   255
         Left            =   3200
         TabIndex        =   70
         Top             =   1580
         Width           =   4300
      End
      Begin VB.Label lbl_depth 
         Caption         =   "Color depth:"
         Height          =   255
         Left            =   3200
         TabIndex        =   77
         Top             =   3175
         Width           =   4300
      End
   End
   Begin VB.Frame fra_tab 
      Height          =   7040
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
         Height          =   6020
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   83
         Top             =   320
         Width           =   7560
      End
      Begin VB.CommandButton btn_info_refresh 
         Caption         =   "Refresh"
         Height          =   345
         Left            =   120
         TabIndex        =   84
         Top             =   6460
         Width           =   1200
      End
      Begin VB.CommandButton btn_info_copy 
         Caption         =   "Copy"
         Height          =   345
         Left            =   1400
         TabIndex        =   85
         Top             =   6460
         Width           =   1200
      End
   End
   Begin VB.CheckBox check_cmd 
      Caption         =   "Run in CMD"
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   8090
      Width           =   1500
   End
   Begin VB.CommandButton btn_preview 
      Caption         =   "Copy commands"
      Height          =   375
      Left            =   120
      TabIndex        =   88
      Top             =   8720
      Width           =   1500
   End
   Begin VB.CommandButton btn_createobjs 
      Caption         =   "Create OBJs"
      Height          =   375
      Left            =   1750
      TabIndex        =   91
      ToolTipText     =   "Runs all steps except the link, then renames every .obj of the source directory to <name>_sav.obj"
      Top             =   8720
      Width           =   1500
   End
   Begin VB.CommandButton btn_ok 
      Caption         =   "Compile"
      Height          =   375
      Left            =   5640
      TabIndex        =   89
      Top             =   8720
      Width           =   1095
   End
   Begin VB.CommandButton btn_exit 
      Caption         =   "Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   90
      Top             =   8720
      Width           =   1095
   End
   Begin VB.Label lbl_cfg 
      Caption         =   "Config:"
      Height          =   255
      Left            =   120
      TabIndex        =   96
      Top             =   165
      Width           =   735
   End
   Begin VB.Label lbl_status 
      Height          =   255
      Left            =   120
      TabIndex        =   87
      Top             =   8390
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
      Begin VB.Menu mnu_tool_manifest 
         Caption         =   "VBMF.exe"
      End
      Begin VB.Menu mnu_tool_runtime 
         Caption         =   "MSVBVM60.DLL"
      End
   End
End
Attribute VB_Name = "frmLinker"
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
'   3. VBStatic  [/DLL:x] [/NOICON] <msvbvm60.dll> <main.obj>   (embed the runtime)
'   4. VBMF      [/ID:n] <main.obj> <manifest>           (embed the application manifest)
'   5. LINK2.EXE
'
' One main.obj (General tab) is shared by all patch steps. The manifest is generated from
' the General tab; its name / version are read from the VERSIONINFO resource of main.obj.
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
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
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
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal dstx As Long, ByVal dsty As Long, ByVal dstw As Long, ByVal dsth As Long, ByVal srcx As Long, ByVal srcy As Long, ByVal srcw As Long, ByVal srch As Long, ByVal srcUnit As Long, ByVal imageAttributes As Long, ByVal callback As Long, ByVal callbackData As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As Long, ByVal wrap As Long, ByVal argb As Long, ByVal bClamp As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long

Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hModule As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hModule As Long, ByVal hResInfo As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Byte, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function PrivateExtractIcons Lib "user32" Alias "PrivateExtractIconsA" (ByVal lpszFile As String, ByVal nIconIndex As Long, ByVal cxIcon As Long, ByVal cyIcon As Long, ByRef phicon As Long, ByRef piconid As Long, ByVal nIcons As Long, ByVal flags As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long

Private Const LOAD_AS_RESOURCE As Long = &H22&       'LOAD_LIBRARY_AS_DATAFILE Or LOAD_LIBRARY_AS_IMAGE_RESOURCE
Private Const RT_ICON As Long = 3
Private Const RT_GROUP_ICON As Long = 14
Private Const RT_VERSION As Long = 16
Private Const DI_NORMAL As Long = 3
Private Const SYS_ICON_ID As Long = 15               'imageres.dll: default exe icon
Private Const VB_ICON_ID As Long = 1                 'main.obj: VB6 default icon
Private Const APP_ICON_NAME As String = "APPICON"
Private Const FLT_EXE As String = "Executables (*.exe)|*.exe|All files (*.*)|*.*"

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
Private Const TOOL_MANIFEST As Long = 4

Private Const FLAG_USER As Long = 0        'ListBox.ItemData: added / edited by the user
Private Const FLAG_CMD As Long = 1         'ListBox.ItemData: came from the command line

'---- state -----------------------------------------------------------------------
Private mLoading As Boolean
Private mDirty As Boolean
Private mExitCode As Long
Private mSecBuf As String                            'key=value<NUL>... block built by SaveConfig

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
Private mToolManifest As String
Private mRuntime As String

Private mFnName() As String                'functions found in the OBJs
Private mFnScope() As String
Private mFnDecor() As String
Private mFnCount As Long
Private mFnStale As Boolean                'function list is read when the DLL tab is opened


Private mGdipToken As Long                 'GDI+ (icon preview)
Private mSrcW As Long                      'size of the icon source image (0 = none, -1 = unknown)
Private mSrcH As Long
Private mSrcCaption As String              'caption / preview path of the selected source image
Private mSrcDraw As String
Private mPrevKey As String                 'state of the last icon preview (skips identical redraws)

Private mScanned As Boolean                'main.obj icon cache (re-read when the file changes)
Private mObjKey As String
Private mHasApp As Boolean                 'main.obj has an APPICON group icon
Private mAppHas As Boolean
Private mIcoApp() As Byte
Private mAppW As Long
Private mAppH As Long
Private mAppBits As Long
Private mObjHas As Boolean                 'main.obj has the VB6 icon (group icon 1)
Private mIcoObj() As Byte
Private mObjW As Long
Private mObjH As Long
Private mObjBits As Long
Private mSysTried As Boolean               'imageres.dll icon cache
Private mSysHas As Boolean
Private mIcoSys() As Byte
Private mSysW As Long
Private mSysH As Long
Private mSysBits As Long
Private mSysHIcon As Long                  'fallback: icon handle from PrivateExtractIcons

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
    cbo_tool.AddItem "vbicon"
    cbo_tool.AddItem "vbmf"

    cbo_uac.AddItem "default"
    cbo_uac.AddItem "asInvoker"
    cbo_uac.AddItem "requireAdministrator"
    cbo_uac.AddItem "highestAvailable"
    cbo_os.AddItem "default"
    cbo_os.AddItem "Windows Vista"
    cbo_os.AddItem "Windows 7"
    cbo_os.AddItem "Windows 8"
    cbo_os.AddItem "Windows 8.1"
    cbo_os.AddItem "Windows 10 / 11"
    cbo_uac.ListIndex = 0
    cbo_os.ListIndex = 0

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
    If mSysHIcon <> 0 Then DestroyIcon mSysHIcon: mSysHIcon = 0
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
    If Index = 1 And mFnStale Then ScanFunctions
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
    Dim ch As Boolean
    Dim mf As Boolean
    Dim i As Long

    s = (check_static.value = vbChecked)
    e = (check_extern.value = vbChecked)
    ic = (check_icon.value = vbChecked)
    mf = (check_manifest.value = vbChecked)
    ScanObjIcon
    rm = IconRemoveOn()
    ch = ic And Not rm And Not mHasApp              'options that change the icon (off when APPICON exists)

    'the one main.obj is needed by every patch step
    lbl_main_obj.Enabled = (s Or e Or ic Or mf)
    txt_main_obj.Enabled = (s Or e Or ic Or mf)
    btn_main_auto.Enabled = (s Or e Or ic Or mf)
    btn_main_browse.Enabled = (s Or e Or ic Or mf)
    lbl_main_state.Enabled = (s Or e Or ic Or mf)

    check_autofix.Enabled = s
    check_noicon.Enabled = s

    'manifest (VBMF)
    check_vstyles.Enabled = mf
    check_dpi.Enabled = mf
    check_segheap.Enabled = mf
    lbl_uac.Enabled = mf
    cbo_uac.Enabled = mf
    lbl_os.Enabled = mf
    cbo_os.Enabled = mf
    lbl_manifest.Enabled = mf
    txt_manifest.Enabled = mf
    btn_manifest_browse.Enabled = mf
    btn_createmanifest.Enabled = mf
    check_automanifest.Enabled = mf

    UpdateArgsEnabled

    'icon tab
    check_remove.Enabled = ic And Not NoObjIcon()   'nothing to remove when main.obj has no icon
    lbl_icon_file.Enabled = ch
    txt_icon_file.Enabled = ch
    btn_icon_pick.Enabled = ch
    btn_icon_clear.Enabled = ch
    check_stretch.Enabled = ch
    lbl_res.Enabled = ch
    lbl_depth.Enabled = ch
    For i = 0 To 4
        check_depth(i).Enabled = ch
    Next i
    UpdateResEnable False
    UpdateIconPreview

    'DLL tab: only for /DLL builds, exports only for standard DLLs
    isDll = IsDllBuild()
    opt_tab(1).Enabled = isDll
    opt_tab(2).Enabled = Not isDll                  'no icon tab for DLLs
    If isDll Then
        If opt_tab(2).value Then opt_tab(0).value = True
    End If
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
        Case TOOL_ICON: en = IconStepOn()
        Case TOOL_MANIFEST: en = (check_manifest.value = vbChecked)
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

Private Sub check_noicon_Click()
    MarkDirty
End Sub

Private Sub check_manifest_Click()
    If mLoading Then Exit Sub
    MarkDirty
    RefreshState
End Sub

Private Sub check_vstyles_Click()
    MarkDirty
End Sub

Private Sub check_dpi_Click()
    MarkDirty
End Sub

Private Sub check_segheap_Click()
    MarkDirty
End Sub

Private Sub check_automanifest_Click()
    MarkDirty
End Sub

Private Sub cbo_uac_Click()
    MarkDirty
End Sub

Private Sub cbo_os_Click()
    MarkDirty
End Sub

Private Sub txt_manifest_Change()
    MarkDirty
End Sub

Private Sub txt_main_obj_LostFocus()
    If mLoading Then Exit Sub
    RefreshState
End Sub

'==============================================================================
' Argument lists (0 = link, 1 = vbstatic, 2 = vbextern)
'==============================================================================
Private Sub cbo_tool_Click()

    Dim i As Long

    If cbo_tool.ListIndex < 0 Then Exit Sub
    For i = 0 To 4
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
    Dim sig As Integer

    On Error GoTo EH
    If Not FileExists(path) Then Exit Function
    n = FileLen(path)
    If n < 20 Then Exit Function

    h = FreeFile
    Open path For Binary Access Read As #h
    Get #h, 1, sig
    If sig <> &H14C Then Close #h: h = 0: Exit Function  'not an i386 COFF object: do not read the whole file
    ReDim mOb(n - 1)
    Get #h, 1, mOb
    Close #h
    h = 0

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

'Zero terminated ANSI string of the loaded object (at most maxLen bytes)
Private Function ObjAsciiZ(ByVal off As Long, ByVal maxLen As Long) As String

    Dim n As Long
    Dim b() As Byte

    If off < 0 Then Exit Function
    If maxLen > UBound(mOb) + 1 - off Then maxLen = UBound(mOb) + 1 - off
    Do While n < maxLen
        If mOb(off + n) = 0 Then Exit Do
        n = n + 1
    Loop
    If n = 0 Then Exit Function
    ReDim b(n - 1)
    CopyMemory b(0), mOb(off), n
    ObjAsciiZ = StrConv(b, vbUnicode)

End Function

Private Function ObjSymName(ByVal i As Long) As String

    Dim p As Long

    p = mObSymOff + i * 18&
    If I32(p) = 0 Then                              'long name: offset into the string table
        ObjSymName = ObjAsciiZ(mObStr + I32(p + 4), &H7FFFFFFF)
    Else
        ObjSymName = ObjAsciiZ(p, 8)
    End If

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

    txt_main_obj.Text = m
    lbl_main_state.Caption = st
    If Not mLoading Then RefreshState

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

Private Sub btn_main_auto_Click()
    AutoDetectObjs
End Sub

Private Function BrowseObj(txt As TextBox, lbl As label) As Boolean

    Dim c As Collection

    Set c = FileDialog("Select main.obj", "Object files (*.obj)|*.obj|All files (*.*)|*.*", False, False, txt.Text)
    If c Is Nothing Then Exit Function
    txt.Text = c(1)
    lbl.Caption = "selected manually"
    BrowseObj = True

End Function

Private Sub btn_main_browse_Click()
    If BrowseObj(txt_main_obj, lbl_main_state) Then RefreshState
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
    Dim cap As Long
    Dim sym As String
    Dim nm As String
    Dim sc As String

    cap = 64
    ReDim mFnName(cap - 1)
    ReDim mFnScope(cap - 1)
    ReDim mFnDecor(cap - 1)
    mFnCount = 0
    mFnStale = False
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
                                If mFnCount >= cap Then                     'grow in big steps (no ReDim per function)
                                    cap = cap * 2
                                    ReDim Preserve mFnName(cap - 1)
                                    ReDim Preserve mFnScope(cap - 1)
                                    ReDim Preserve mFnDecor(cap - 1)
                                End If
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

    m = MainObj()
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

    Dim msg As String

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
    msg = Err.Description                            'read first: any On Error statement clears Err
    On Error Resume Next
    Close #f
    MsgBox "Cannot write " & p & vbCrLf & msg, vbExclamation

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

    a = SplitParams(c)
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
        Case TOOL_ICON: ListKey = "IconArgs"
        Case Else: ListKey = "ManifestArgs"
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
    For k = 0 To 4
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
    check_noicon.value = vbUnchecked
    check_manifest.value = vbUnchecked
    check_vstyles.value = vbUnchecked
    check_dpi.value = vbUnchecked
    check_segheap.value = vbUnchecked
    check_automanifest.value = vbChecked
    cbo_uac.ListIndex = 0
    cbo_os.ListIndex = 0
    For i = 0 To 5
        check_res(i).value = vbChecked
    Next i
    For i = 0 To 4
        check_depth(i).value = IIf(i = 4, vbChecked, vbUnchecked)
    Next i
    opt_dll(0).value = True
    lst_exports.Clear
    txt_def.Text = ""
    txt_manifest.Text = ""
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
        check_noicon.value = BoolCheck(IniLong(f, sec, "StaticNoIcon", 0))
        check_manifest.value = BoolCheck(IniLong(f, sec, "Manifest", 0))
        check_vstyles.value = BoolCheck(IniLong(f, sec, "ManifestVisualStyles", 0))
        check_dpi.value = BoolCheck(IniLong(f, sec, "ManifestDpiAware", 0))
        check_segheap.value = BoolCheck(IniLong(f, sec, "ManifestSegmentHeap", 0))
        check_automanifest.value = BoolCheck(IniLong(f, sec, "AutoManifest", 1))
        SelectCombo cbo_uac, IniGet(f, sec, "ManifestUAC", "default")
        SelectCombo cbo_os, IniGet(f, sec, "ManifestOS", "default")
        txt_manifest.Text = IniGet(f, sec, "ManifestFile", "")
        txt_icon_file.Text = Dec(IniGet(f, sec, "IconFile", ""))

        lst = IniGet(f, sec, "IconRes", "*")
        If lst <> "*" Then
            For i = 0 To 5
                check_res(i).value = IIf(InList(lst, ResSize(i)), vbChecked, vbUnchecked)
            Next i
        End If
        lst = IniGet(f, sec, "IconDepth", "*")
        If lst <> "*" Then
            For i = 0 To 4
                check_depth(i).value = IIf(InList(lst, DepthBits(i)), vbChecked, vbUnchecked)
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

        For k = 0 To 4
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
    If Len(Trim$(txt_manifest.Text)) = 0 Then txt_manifest.Text = DefaultManifestPath()
    If fra_tab(1).Visible Then ScanFunctions Else mFnStale = True
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

    mSecBuf = "Version=2" & vbNullChar                     'whole section is written in one go below

    SecAdd "Static", IIf(check_static.value = vbChecked, "1", "0")
    SecAdd "Extern", IIf(check_extern.value = vbChecked, "1", "0")
    SecAdd "Icon", IIf(check_icon.value = vbChecked, "1", "0")
    SecAdd "IconRemove", IIf(check_remove.value = vbChecked, "1", "0")
    SecAdd "IconStretch", IIf(check_stretch.value = vbChecked, "1", "0")
    SecAdd "IconFile", Enc(Trim$(txt_icon_file.Text))
    SecAdd "IconRes", SavedResList()
    SecAdd "IconDepth", SavedDepthList()
    SecAdd "AutoFix", IIf(check_autofix.value = vbChecked, "1", "0")
    SecAdd "RunInCmd", IIf(check_cmd.value = vbChecked, "1", "0")
    SecAdd "AutoDef", IIf(check_autodef.value = vbChecked, "1", "0")
    SecAdd "DllMode", IIf(opt_dll(1).value, "1", "0")
    SecAdd "DefFile", Trim$(txt_def.Text)
    SecAdd "StaticNoIcon", IIf(check_noicon.value = vbChecked, "1", "0")
    SecAdd "Manifest", IIf(check_manifest.value = vbChecked, "1", "0")
    SecAdd "ManifestVisualStyles", IIf(check_vstyles.value = vbChecked, "1", "0")
    SecAdd "ManifestDpiAware", IIf(check_dpi.value = vbChecked, "1", "0")
    SecAdd "ManifestSegmentHeap", IIf(check_segheap.value = vbChecked, "1", "0")
    SecAdd "ManifestUAC", ComboText(cbo_uac)
    SecAdd "ManifestOS", ComboText(cbo_os)
    SecAdd "ManifestFile", Trim$(txt_manifest.Text)
    SecAdd "AutoManifest", IIf(check_automanifest.value = vbChecked, "1", "0")

    'argument lists: user entries in full, command line entries only when unchecked
    For k = 0 To 4
        n = 0
        For i = 0 To args_list(k).ListCount - 1
            kind = ""
            If args_list(k).ItemData(i) = FLAG_CMD Then
                If Not args_list(k).Selected(i) Then kind = "off"
            Else
                kind = "user"
            End If
            If Len(kind) > 0 Then
                SecAdd ListKey(k) & "." & n, Enc(args_list(k).List(i))
                SecAdd ListKey(k) & "." & n & ".Kind", kind
                SecAdd ListKey(k) & "." & n & ".On", IIf(args_list(k).Selected(i), "1", "0")
                n = n + 1
            End If
        Next i
        SecAdd ListKey(k) & ".Count", CStr(n)
    Next k

    SecAdd "Hide.Count", CStr(mHiddenCount)
    For i = 0 To mHiddenCount - 1
        SecAdd "Hide." & i, Enc(mHidden(i))
    Next i

    SecAdd "Exports.Count", CStr(lst_exports.ListCount)
    For i = 0 To lst_exports.ListCount - 1
        SecAdd "Exports." & i, Enc(lst_exports.List(i))
    Next i

    If Not FlushSection(sec, f) Then
        MsgBox "Cannot write the config file:" & vbCrLf & f, vbExclamation
        Exit Function
    End If
    IniPut f, "VBLinker", "Active", cfgName
    SaveConfig = True

End Function

Private Sub SecAdd(ByVal key As String, ByVal txt As String)

    mSecBuf = mSecBuf & key & "=" & txt & vbNullChar

End Sub

'Replaces the whole section with mSecBuf. The API takes at most 32767 characters per section,
'bigger blocks are written key by key.
Private Function FlushSection(ByVal sec As String, ByVal f As String) As Boolean

    Dim parts() As String
    Dim i As Long
    Dim p As Long

    If Len(mSecBuf) < 30000 Then
        FlushSection = (WritePrivateProfileSection(sec, mSecBuf & vbNullChar, f) <> 0)
        Exit Function
    End If

    WritePrivateProfileString sec, vbNullString, vbNullString, f
    FlushSection = True
    parts = Split(mSecBuf, vbNullChar)
    For i = 0 To UBound(parts)
        p = InStr(parts(i), "=")
        If p > 1 Then
            If WritePrivateProfileString(sec, Left$(parts(i), p - 1), Mid$(parts(i), p + 1), f) = 0 Then FlushSection = False
        End If
    Next i

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
    mToolManifest = IniGet(AppIni(), "Tools", "VBMF", AddBS(App.path) & "VBMF.exe")
    mRuntime = IniGet(AppIni(), "Tools", "Runtime", DefaultRuntime())
End Sub

Private Sub SaveTools()
    IniPut AppIni(), "Tools", "Link2", mToolLink
    IniPut AppIni(), "Tools", "VBStatic", mToolStatic
    IniPut AppIni(), "Tools", "VBExtern", mToolExtern
    IniPut AppIni(), "Tools", "VBIcon", mToolIcon
    IniPut AppIni(), "Tools", "VBMF", mToolManifest
    IniPut AppIni(), "Tools", "Runtime", mRuntime
End Sub

Private Sub btn_tools_Click()
    mnu_tool_link.Caption = "LINK2.EXE: " & Replace(mToolLink, "&", "&&")
    mnu_tool_static.Caption = "VBStatic.exe: " & Replace(mToolStatic, "&", "&&")
    mnu_tool_extern.Caption = "VBExtern.exe: " & Replace(mToolExtern, "&", "&&")
    mnu_tool_icon.Caption = "VBIcon.exe: " & Replace(mToolIcon, "&", "&&")
    mnu_tool_manifest.Caption = "VBMF.exe: " & Replace(mToolManifest, "&", "&&")
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
    PickTool mToolLink, "Select the real linker (LINK2.EXE)", FLT_EXE
End Sub

Private Sub mnu_tool_static_Click()
    PickTool mToolStatic, "Select VBStatic.exe", FLT_EXE
End Sub

Private Sub mnu_tool_extern_Click()
    PickTool mToolExtern, "Select VBExtern.exe", FLT_EXE
End Sub

Private Sub mnu_tool_icon_Click()
    PickTool mToolIcon, "Select VBIcon.exe", FLT_EXE
End Sub

Private Sub mnu_tool_manifest_Click()
    PickTool mToolManifest, "Select VBMF.exe", FLT_EXE
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
    Dim msg As String

    If Not PrepareBuild(True) Then Exit Sub

    Me.Enabled = False
    On Error GoTo EH
    rc = RunPipeline(False, rep)
    On Error GoTo 0
    Me.Enabled = True

    If rc = 0 Then
        mExitCode = 0
        Unload Me
    Else
        mExitCode = rc
    End If
    Exit Sub

EH:
    msg = Err.Description
    Me.Enabled = True
    mExitCode = 1
    MsgBox "The build was aborted:" & vbCrLf & msg, vbCritical

End Sub

'Create OBJs: every step except the link, then every .obj of the source directory is renamed to
'<name>_sav.obj so that VB6 cannot delete them after the build.
Private Sub btn_createobjs_Click()

    Dim rc As Long
    Dim rep As String
    Dim dirPath As String
    Dim n As Long
    Dim failed As String
    Dim msg As String

    dirPath = SourceDir()
    If Len(dirPath) = 0 Then
        MsgBox "The source directory is unknown: select main.obj (General tab) or check an .obj in the link list.", vbExclamation
        Exit Sub
    End If

    If Not PrepareBuild(False) Then Exit Sub

    Me.Enabled = False
    On Error GoTo EH
    rc = RunPipeline(False, rep, True)
    If rc = 0 Then n = SaveObjs(dirPath, failed)
    On Error GoTo 0
    Me.Enabled = True

    If rc <> 0 Then
        mExitCode = rc
        Exit Sub
    End If

    msg = n & " OBJ file(s) renamed to <name>_sav.obj in:" & vbCrLf & dirPath
    If Len(failed) > 0 Then msg = msg & vbCrLf & vbCrLf & "Could not rename:" & vbCrLf & failed
    MsgBox msg, IIf(Len(failed) > 0, vbExclamation, vbInformation)

    mExitCode = 1                                    'nothing was linked: no executable exists
    Unload Me
    Exit Sub

EH:
    msg = Err.Description
    Me.Enabled = True
    mExitCode = 1
    MsgBox "The operation was aborted:" & vbCrLf & msg, vbCritical

End Sub

'checks + automatic manifest / .DEF that precede the pipeline (Compile and Create OBJs)
Private Function PrepareBuild(ByVal needLink As Boolean) As Boolean

    If Not CheckSetup(needLink) Then Exit Function

    'the manifest is built from the project info of the untouched main.obj
    If check_manifest.value = vbChecked And check_automanifest.value = vbChecked Then
        If Not CreateManifest() Then Exit Function
    End If

    If IsDllBuild() And opt_dll(1).value And check_autodef.value = vbChecked And lst_exports.ListCount > 0 Then
        If Not CreateDef(True) Then Exit Function
    End If

    PrepareBuild = True

End Function

'directory of main.obj (or of the first checked .obj of the link list)
Private Function SourceDir() As String

    Dim m As String
    Dim v As Variant

    m = MainObj()
    If Len(m) = 0 Then
        For Each v In ObjsFromList()
            m = CStr(v)
            Exit For
        Next v
    End If
    If Len(m) = 0 Then Exit Function

    SourceDir = DirOf(m)
    If Len(SourceDir) = 0 Then SourceDir = CurDir$

End Function

'Renames every *.obj of dirPath to <name>_sav.obj (an existing _sav file is replaced, files that
'already end in _sav are left alone). Returns the number of renamed files, failures go to failed.
Private Function SaveObjs(ByVal dirPath As String, ByRef failed As String) As Long

    Dim names As New Collection
    Dim v As Variant
    Dim f As String
    Dim stem As String
    Dim src As String
    Dim dst As String

    dirPath = AddBS(dirPath)

    f = Dir$(dirPath & "*.obj")
    Do While Len(f) > 0
        If LCase$(Right$(f, 4)) = ".obj" Then
            If LCase$(Right$(Left$(f, Len(f) - 4), 4)) <> "_sav" Then names.Add f
        End If
        f = Dir$
    Loop

    For Each v In names
        f = CStr(v)
        stem = Left$(f, Len(f) - 4)
        src = dirPath & f
        dst = dirPath & stem & "_sav.obj"
        On Error Resume Next
        Err.Clear
        If FileExists(dst) Then Kill dst
        If Err.Number = 0 Then Name src As dst
        If Err.Number <> 0 Then
            failed = failed & f & " (" & Err.Description & ")" & vbCrLf
        Else
            SaveObjs = SaveObjs + 1
        End If
        Err.Clear
        On Error GoTo 0
    Next v

End Function

Private Function InLinkList(ByVal path As String) As Boolean

    Dim c As Collection
    Dim v As Variant

    Set c = ObjsFromList()
    If c.count = 0 Then InLinkList = True: Exit Function     'no .obj arguments at all: nothing to compare
    For Each v In c
        If StrComp(CStr(v), path, vbTextCompare) = 0 Then InLinkList = True: Exit Function
    Next v

End Function

Private Function CheckSetup(Optional ByVal needLink As Boolean = True) As Boolean

    Dim s As Boolean
    Dim e As Boolean
    Dim ic As Boolean
    Dim mf As Boolean
    Dim isDll As Boolean

    s = (check_static.value = vbChecked)
    e = (check_extern.value = vbChecked)
    mf = (check_manifest.value = vbChecked)
    isDll = IsDllBuild()
    ic = (check_icon.value = vbChecked) And Not isDll

    If needLink And Not FileExists(mToolLink) Then
        MsgBox "Linker not found:" & vbCrLf & mToolLink & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
        Exit Function
    End If

    If s Or e Or ic Or mf Then
        If Not FileExists(MainObj()) Then
            MsgBox "main.obj was not found." & vbCrLf & "Use Auto or select it manually (General tab).", vbExclamation
            Exit Function
        End If
        If needLink Then
            If Not InLinkList(MainObj()) Then
                If MsgBox("main.obj is not one of the checked .obj arguments of the link list:" & vbCrLf & MainObj() & vbCrLf & vbCrLf & _
                          "The patched file would not be linked. Continue anyway?", vbQuestion + vbYesNo) <> vbYes Then Exit Function
            End If
        End If
    End If

    If e Then
        If Not FileExists(mToolExtern) Then
            MsgBox "VBExtern.exe not found:" & vbCrLf & mToolExtern & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
    End If

    If ic Then
        If Not FileExists(mToolIcon) Then
            MsgBox "VBIcon.exe not found:" & vbCrLf & mToolIcon & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If Not EffectiveRemove() Then
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
    End If

    If mf Then
        If Not FileExists(mToolManifest) Then
            MsgBox "VBMF.exe not found:" & vbCrLf & mToolManifest & vbCrLf & vbCrLf & "Set it under Tools...", vbExclamation
            Exit Function
        End If
        If Len(Trim$(txt_manifest.Text)) = 0 And check_automanifest.value <> vbChecked Then
            MsgBox "VBMF: please enter the manifest file (General tab).", vbExclamation
            Exit Function
        End If
        If check_automanifest.value <> vbChecked And Not FileExists(Trim$(txt_manifest.Text)) Then
            MsgBox "VBMF: the manifest file was not found:" & vbCrLf & Trim$(txt_manifest.Text) & vbCrLf & vbCrLf & _
                   "Use 'Create manifest' or enable the automatic regeneration.", vbExclamation
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


'Runs one pipeline step (with dryRun only adds its command line to the report). Returns the exit code.
Private Function RunTool(ByVal tool As String, ByVal src As String, ByVal cmd As String, ByVal isLast As Boolean, ByVal dryRun As Boolean, ByRef report As String) As Long

    If dryRun Then
        report = report & cmd & vbCrLf
    ElseIf Len(src) > 0 Then
        RunTool = RunStep(tool & " - " & FileOf(src), cmd, isLast)
    Else
        RunTool = RunStep(tool, cmd, isLast)
    End If

End Function

'Runs (or, with dryRun, only lists) VBExtern -> VBIcon -> VBStatic -> VBMF -> LINK2. Returns 0 on success.
'With skipLink the last step (LINK2) is left out. The tools patch main.obj in place; nothing is
'restored when a step fails.
Private Function RunPipeline(ByVal dryRun As Boolean, ByRef report As String, Optional ByVal skipLink As Boolean = False) As Long

    Dim isDll As Boolean
    Dim useStatic As Boolean
    Dim useExtern As Boolean
    Dim useIcon As Boolean
    Dim useManifest As Boolean
    Dim src As String
    Dim mf As String
    Dim cmd As String
    Dim rc As Long
    Dim result As Long

    isDll = IsDllBuild()
    useStatic = (check_static.value = vbChecked)
    useExtern = (check_extern.value = vbChecked)
    useIcon = (check_icon.value = vbChecked) And Not isDll
    useManifest = (check_manifest.value = vbChecked)
    src = MainObj()                                  'one main.obj for all patch steps
    report = ""

    '---- 1. VBExtern ---------------------------------------------------------
    If useExtern Then
        cmd = q(mToolExtern) & ArgsFor(TOOL_EXTERN, True) & " " & q(src) & ArgsFor(TOOL_EXTERN, False)
        rc = RunTool("VBExtern", src, cmd, False, dryRun, report)
        If rc <> 0 Then StepFailed "VBExtern", rc: result = rc: GoTo Done
    End If

    '---- 2. VBIcon -----------------------------------------------------------
    If useIcon Then
        cmd = q(mToolIcon) & ArgsFor(TOOL_ICON, True) & " " & q(src)
        If EffectiveRemove() Then
            cmd = cmd & " /REMOVE"
        Else
            cmd = cmd & " " & q(Trim$(txt_icon_file.Text)) & " /RES:" & ResList() & " /DEPTH:" & DepthList()
            If check_stretch.value = vbChecked Then cmd = cmd & " /STRETCH"
        End If
        cmd = cmd & ArgsFor(TOOL_ICON, False)
        rc = RunTool("VBIcon", src, cmd, False, dryRun, report)
        If rc <> 0 Then StepFailed "VBIcon", rc: result = rc: GoTo Done
    End If

    '---- 3. VBStatic ---------------------------------------------------------
    If useStatic Then
        cmd = q(mToolStatic) & ArgsFor(TOOL_STATIC, True)
        If isDll Then
            If opt_dll(1).value Then cmd = cmd & " /DLL:STANDARD" Else cmd = cmd & " /DLL:ACTIVEX"
        End If
        If check_noicon.value = vbChecked And Not ToolHasArg(TOOL_STATIC, "/NOICON") Then cmd = cmd & " /NOICON"
        cmd = cmd & " " & q(mRuntime) & " " & q(src) & ArgsFor(TOOL_STATIC, False)
        rc = RunTool("VBStatic", src, cmd, False, dryRun, report)
        If rc <> 0 Then StepFailed "VBStatic", rc: result = rc: GoTo Done
    End If

    '---- 4. VBMF (after VBStatic: replaces a manifest of the same ID, e.g. one merged from the runtime) ----
    If useManifest Then
        mf = Trim$(txt_manifest.Text)
        cmd = q(mToolManifest) & ArgsFor(TOOL_MANIFEST, True)
        If isDll And Len(ToolArgValue(TOOL_MANIFEST, "/ID:")) = 0 Then cmd = cmd & " /ID:2"   'DLL manifest resource ID
        cmd = cmd & " " & q(src) & " " & q(mf) & ArgsFor(TOOL_MANIFEST, False)
        rc = RunTool("VBMF", src, cmd, False, dryRun, report)
        If rc <> 0 Then StepFailed "VBMF", rc: result = rc: GoTo Done
    End If

    '---- 5. link -------------------------------------------------------------
    If Not skipLink Then
        cmd = q(mToolLink) & BuildLinkArgs(isDll, useStatic)
        rc = RunTool("LINK2", "", cmd, True, dryRun, report)
        If rc <> 0 Then StepFailed "LINK2", rc: result = rc
    End If

Done:
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
    MsgBox msg, vbExclamation

End Sub

Private Function NormSw(ByVal t As String) As String

    If Left$(t, 1) = "-" Then NormSw = "/" & Mid$(t, 2) Else NormSw = t

End Function

'Switches (/xxx) or positional arguments of the vbstatic / vbextern list
Private Function ArgsFor(ByVal tool As Long, ByVal switches As Boolean) As String

    Dim i As Long
    Dim t As String
    Dim c As String
    Dim s As String

    For i = 0 To args_list(tool).ListCount - 1
        If args_list(tool).Selected(i) Then
            t = args_list(tool).List(i)
            If tool <> TOOL_LINK Then t = NormSw(t)       'the tools only know "/" (link.exe takes both)
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
    Dim hasEntry As Boolean
    Dim hasNoDef As Boolean

    doFix = useStatic And (check_autofix.value = vbChecked)
    If isDll Then entry = "_VBStaticDllEntry@12" Else entry = "_VBStaticEntry"

    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then
            t = args_list(TOOL_LINK).List(i)
            u = UCase$(StripQ(t))

            If doFix Then
                If Left$(u, 7) = "/ENTRY:" Or Left$(u, 7) = "-ENTRY:" Then
                    t = "/ENTRY:" & entry
                    hasEntry = True
                ElseIf u = "/NODEFAULTLIB" Or u = "-NODEFAULTLIB" Then
                    hasNoDef = True
                End If
            End If

            s = s & " " & t
        End If
    Next i

    If doFix Then
        If Not hasEntry Then s = s & " /ENTRY:" & entry
        If Not hasNoDef Then s = s & " /NODEFAULTLIB"
    End If
    BuildLinkArgs = s

End Function

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

    initPath = Left$(StripQ(initPath), 2000)        'the 4096 byte buffer must hold it
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
Private Function q(ByVal s As String) As String
    q = """" & s & """"
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


'==============================================================================
' COFF helpers for the loaded object (sections / counters)
'==============================================================================
Private Function ObjSecName(ByVal i As Long) As String
    ObjSecName = ObjAsciiZ(20 + 40& * i, 8)
End Function

Private Function ObjSecText(ByVal i As Long) As String

    Dim hd As Long
    Dim raw As Long
    Dim sz As Long
    Dim b() As Byte

    hd = 20 + 40& * i
    raw = I32(hd + 20)
    sz = I32(hd + 16)
    If raw <= 0 Or sz <= 0 Then Exit Function
    If raw + sz > UBound(mOb) + 1 Then Exit Function
    ReDim b(sz - 1)
    CopyMemory b(0), mOb(raw), sz
    ObjSecText = StrConv(b, vbUnicode)

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
'==============================================================================
' Icon resources: main.obj (.rsrc$01 / .rsrc$02) and imageres.dll
'==============================================================================
Private Function BU16(a() As Byte, ByVal p As Long) As Long
    BU16 = CLng(a(p)) + CLng(a(p + 1)) * 256&
End Function

'section of the loaded COFF object by name (raw data offset / size)
Private Function ObjSection(ByVal secName As String, ByRef off As Long, ByRef size As Long) As Boolean
    Dim idx As Long
    Dim relOff As Long
    Dim relCnt As Long
    ObjSection = ObjSectionEx(secName, off, size, idx, relOff, relCnt)
End Function

'section of the loaded COFF object by name: raw data offset / size, 1-based index, relocation table
Private Function ObjSectionEx(ByVal secName As String, ByRef off As Long, ByRef size As Long, ByRef idx As Long, ByRef relOff As Long, ByRef relCnt As Long) As Boolean

    Dim n As Long
    Dim p As Long
    Dim i As Long

    On Error GoTo EH
    n = U16(2)
    p = 20& + U16(16)
    For i = 0 To n - 1
        If p + 40 > UBound(mOb) + 1 Then Exit For
        If ObjAsciiZ(p, 8) = secName Then
            size = I32(p + 16)
            off = I32(p + 20)
            relOff = I32(p + 24)
            relCnt = U16(p + 32)
            idx = i + 1
            If off > 0 And size > 0 Then
                If off + size <= UBound(mOb) + 1 Then ObjSectionEx = True
            End If
            Exit Function
        End If
        p = p + 40
    Next i
EH:

End Function

'Offset (inside .rsrc$02) of the data of the resource data entry at "entry" (inside .rsrc$01).
'The OffsetToData field of an OBJ is only the addend: the real offset is the value of the symbol
'that the relocation of that field points to.
Private Function ResDataOffset(ByVal b1 As Long, ByVal entry As Long, ByVal relOff As Long, ByVal relCnt As Long) As Long

    Dim i As Long
    Dim q As Long
    Dim sp As Long

    On Error GoTo EH
    ResDataOffset = I32(b1 + entry)
    For i = 0 To relCnt - 1
        q = relOff + i * 10&
        If I32(q) = entry Then
            sp = mObSymOff + I32(q + 4) * 18&
            ResDataOffset = ResDataOffset + I32(sp + 8)
            Exit Function
        End If
    Next i
EH:
End Function

'Looks up one entry of a resource directory (by id, or by name when nm is given)
Private Function ResDirFind(ByVal base As Long, ByVal dirOff As Long, ByVal id As Long, ByVal nm As String, ByRef target As Long, ByRef isDir As Boolean) As Boolean

    Dim n As Long
    Dim i As Long
    Dim k As Long
    Dim p As Long
    Dim v As Long
    Dim w As Long
    Dim nOff As Long
    Dim ln As Long
    Dim s As String
    Dim hit As Boolean

    On Error GoTo EH
    p = base + dirOff
    n = U16(p + 12) + U16(p + 14)
    For i = 0 To n - 1
        p = base + dirOff + 16 + i * 8
        v = I32(p)
        w = I32(p + 4)
        hit = False
        If Len(nm) > 0 Then
            If v < 0 Then
                nOff = base + (v And &H7FFFFFFF)
                ln = U16(nOff)
                s = ""
                For k = 0 To ln - 1
                    s = s & Chr$(mOb(nOff + 2 + k * 2))
                Next k
                hit = (UCase$(s) = UCase$(nm))
            End If
        ElseIf v >= 0 Then
            hit = (v = id)
        End If
        If hit Then
            target = w And &H7FFFFFFF
            isDir = (w < 0)
            ResDirFind = True
            Exit Function
        End If
    Next i
EH:
End Function

'Reads one resource (type, id or name) of the loaded main.obj
Private Function ObjResFind(ByVal resType As Long, ByVal id As Long, ByVal nm As String, data() As Byte) As Boolean

    Dim b1 As Long
    Dim s1 As Long
    Dim b2 As Long
    Dim s2 As Long
    Dim t As Long
    Dim d As Boolean
    Dim p As Long
    Dim off As Long
    Dim sz As Long
    Dim ix As Long
    Dim relOff As Long
    Dim relCnt As Long

    On Error GoTo EH
    If Not ObjSectionEx(".rsrc$01", b1, s1, ix, relOff, relCnt) Then Exit Function
    If Not ObjSection(".rsrc$02", b2, s2) Then Exit Function
    If Not ResDirFind(b1, 0, resType, "", t, d) Then Exit Function
    If Not d Then Exit Function
    If Not ResDirFind(b1, t, id, nm, t, d) Then Exit Function
    If d Then                                       'language level: first entry
        p = b1 + t
        If U16(p + 12) + U16(p + 14) < 1 Then Exit Function
        If I32(p + 20) < 0 Then Exit Function
        t = I32(p + 20)
    End If
    p = b1 + t
    off = ResDataOffset(b1, t, relOff, relCnt)      'offset inside .rsrc$02
    sz = I32(p + 4)
    If off < 0 Or sz <= 0 Or off + sz > s2 Then Exit Function
    ReDim data(sz - 1)
    CopyMemory data(0), mOb(b2 + off), sz
    ObjResFind = True
EH:
End Function

'Best variant of a group icon (RT_GROUP_ICON data): highest resolution, then highest color depth
Private Function GrpPick(g() As Byte, ByRef id As Long, ByRef w As Long, ByRef h As Long, ByRef bits As Long) As Boolean

    Dim n As Long
    Dim i As Long
    Dim p As Long
    Dim ew As Long
    Dim EH As Long
    Dim eb As Long
    Dim sc As Long
    Dim best As Long

    On Error GoTo EH
    If UBound(g) < 5 Then Exit Function
    n = BU16(g, 4)
    best = -1
    For i = 0 To n - 1
        p = 6 + i * 14
        If p + 13 > UBound(g) Then Exit For
        ew = g(p)
        If ew = 0 Then ew = 256
        EH = g(p + 1)
        If EH = 0 Then EH = 256
        eb = BU16(g, p + 6)
        sc = ew * EH * 100& + eb
        If sc > best Then
            best = sc
            id = BU16(g, p + 12)
            w = ew
            h = EH
            bits = eb
        End If
    Next i
    GrpPick = (best >= 0)
EH:
End Function

'Real size / depth of an icon image (RT_ICON data). The group directory can be wrong
'(e.g. the doubled DIB height), so the image header is the reference.
Private Sub IcoFix(ico() As Byte, ByRef w As Long, ByRef h As Long, ByRef bits As Long)

    Dim nw As Long
    Dim nh As Long
    Dim nb As Long

    On Error GoTo EH
    If UBound(ico) < 23 Then Exit Sub
    If ico(0) = &H89 And ico(1) = &H50 Then         'PNG: big-endian size inside IHDR
        nw = ico(18) * 256& + ico(19)
        nh = ico(22) * 256& + ico(23)
    ElseIf BU16(ico, 0) = 40 Then                   'DIB: the height includes the AND mask
        nw = BU16(ico, 4)
        nh = BU16(ico, 8) \ 2
        nb = BU16(ico, 14)
    End If
    If nw > 0 And nh > 0 Then
        w = nw
        h = nh
        If nb > 0 Then bits = nb
    End If
EH:

End Sub

'The best icon image of a group icon of the loaded main.obj
Private Function ObjIconBest(ByVal nm As String, ByVal id As Long, ico() As Byte, ByRef w As Long, ByRef h As Long, ByRef bits As Long) As Boolean

    Dim g() As Byte
    Dim iid As Long

    If Not ObjResFind(RT_GROUP_ICON, id, nm, g) Then Exit Function
    If Not GrpPick(g, iid, w, h, bits) Then Exit Function
    ObjIconBest = ObjResFind(RT_ICON, iid, "", ico)
    If ObjIconBest Then IcoFix ico, w, h, bits

End Function

'Reads icon 15 of imageres.dll (best variant)
Private Function SysIconLoad() As Boolean

    Dim hm As Long
    Dim hr As Long
    Dim p As Long
    Dim sz As Long
    Dim g() As Byte
    Dim iid As Long

    If mSysTried Then SysIconLoad = mSysHas: Exit Function
    mSysTried = True

    On Error GoTo EH
    hm = LoadLibraryEx(Environ$("windir") & "\System32\imageres.dll", 0, LOAD_AS_RESOURCE)
    If hm = 0 Then Exit Function

    hr = FindResource(hm, SYS_ICON_ID, RT_GROUP_ICON)
    If hr = 0 Then GoTo Done
    sz = SizeofResource(hm, hr)
    p = LockResource(LoadResource(hm, hr))
    If sz < 6 Or p = 0 Then GoTo Done
    ReDim g(sz - 1)
    CopyMemory g(0), ByVal p, sz
    If Not GrpPick(g, iid, mSysW, mSysH, mSysBits) Then GoTo Done

    hr = FindResource(hm, iid, RT_ICON)
    If hr = 0 Then GoTo Done
    sz = SizeofResource(hm, hr)
    p = LockResource(LoadResource(hm, hr))
    If sz < 1 Or p = 0 Then GoTo Done
    ReDim mIcoSys(sz - 1)
    CopyMemory mIcoSys(0), ByVal p, sz
    IcoFix mIcoSys, mSysW, mSysH, mSysBits
    mSysHas = True

Done:
    FreeLibrary hm
    SysIconLoad = mSysHas
    Exit Function
EH:
    If hm <> 0 Then FreeLibrary hm
    mSysHas = False

End Function

'(Re)reads the icons of main.obj when the file changed
Private Sub ScanObjIcon()

    Dim path As String
    Dim key As String
    Dim tmp() As Byte

    On Error Resume Next
    path = MainObj()
    If Len(path) > 0 Then
        If FileExists(path) Then key = path & "|" & FileLen(path) & "|" & FileDateTime(path)
    End If
    If mScanned And key = mObjKey Then Exit Sub
    mScanned = True
    mObjKey = key
    mHasApp = False
    mAppHas = False
    mObjHas = False
    If Len(key) = 0 Then Exit Sub
    On Error GoTo 0

    If Not ObjLoad(path) Then Exit Sub
    mObjHas = ObjIconBest("", VB_ICON_ID, mIcoObj, mObjW, mObjH, mObjBits)
    If ObjResFind(RT_GROUP_ICON, 0, APP_ICON_NAME, tmp) Then
        mHasApp = True
        mAppHas = ObjIconBest(APP_ICON_NAME, 0, mIcoApp, mAppW, mAppH, mAppBits)
    End If

End Sub

'main.obj holds no icon at all (valid after ScanObjIcon): there is nothing to remove
Private Function NoObjIcon() As Boolean
    NoObjIcon = (Len(mObjKey) > 0) And Not mHasApp And Not mObjHas
End Function

'"Remove icon" is ticked and possible (valid after ScanObjIcon)
Private Function IconRemoveOn() As Boolean
    IconRemoveOn = (check_remove.value = vbChecked) And Not NoObjIcon()
End Function

'The icon step runs (never for DLLs)
Private Function IconStepOn() As Boolean
    IconStepOn = (check_icon.value = vbChecked) And Not IsDllBuild()
End Function

'Remove the VB6 icon without adding a new one (explicitly chosen, or APPICON already is the icon)
Private Function EffectiveRemove() As Boolean
    ScanObjIcon
    EffectiveRemove = IconRemoveOn() Or mHasApp
End Function

Private Function IconInfo(ByVal w As Long, ByVal h As Long, ByVal bits As Long) As String
    IconInfo = w & " x " & h & " px, " & bits & " bit"
End Function

'Size (dw x dh) of a w x h image scaled to fit the preview box, aspect ratio kept
Private Sub FitPreview(ByVal w As Long, ByVal h As Long, ByRef dw As Long, ByRef dh As Long)

    Dim sc As Double

    sc = pic_icon.ScaleWidth / w
    If pic_icon.ScaleHeight / h < sc Then sc = pic_icon.ScaleHeight / h
    dw = CLng(w * sc)
    dh = CLng(h * sc)
    If dw < 1 Then dw = 1
    If dh < 1 Then dh = 1

End Sub

Private Sub DrawHIcon(ByVal hIc As Long, ByVal w As Long, ByVal h As Long)

    Dim dw As Long
    Dim dh As Long

    FitPreview w, h, dw, dh
    DrawIconEx pic_icon.hdc, (pic_icon.ScaleWidth - dw) \ 2, (pic_icon.ScaleHeight - dh) \ 2, hIc, dw, dh, 0, 0, DI_NORMAL

End Sub

'Draws an icon image (DIB / PNG data of an RT_ICON resource); False when Windows refuses it
Private Function DrawIconBytes(ico() As Byte, ByVal w As Long, ByVal h As Long) As Boolean

    Dim hIc As Long

    DrawChecker
    On Error Resume Next
    hIc = CreateIconFromResourceEx(ico(0), UBound(ico) + 1, 1, &H30000, w, h, 0)
    If hIc = 0 Then hIc = CreateIconFromResourceEx(ico(0), UBound(ico) + 1, 1, &H30000, 0, 0, 0)
    If hIc <> 0 Then
        DrawHIcon hIc, w, h
        DestroyIcon hIc
        DrawIconBytes = True
    End If
    pic_icon.Refresh

End Function

'Draws the default exe icon of Windows (imageres.dll #15); False when it cannot be loaded
Private Function DrawSysIcon() As Boolean

    Dim id As Long

    If SysIconLoad() Then
        If DrawIconBytes(mIcoSys, mSysW, mSysH) Then DrawSysIcon = True: Exit Function
    Else
        DrawChecker
    End If

    'fallback: let Windows pick the icon (largest size)
    On Error Resume Next
    If mSysHIcon = 0 Then
        PrivateExtractIcons Environ$("windir") & "\System32\imageres.dll", -SYS_ICON_ID, 256, 256, mSysHIcon, id, 1, 0
    End If
    If mSysHIcon <> 0 Then
        DrawHIcon mSysHIcon, 256, 256
        pic_icon.Refresh
        DrawSysIcon = True
    End If

End Function

'Icon preview + the text below it:
'  APPICON in main.obj     -> APPICON (always)
'  Remove icon             -> default exe icon of Windows (imageres.dll #15)
'  source image selected   -> that image
'  otherwise               -> the VB6 icon stored in main.obj (#1)
Private Sub UpdateIconPreview()

    Dim rm As Boolean
    Dim ok As Boolean
    Dim key As String

    ScanObjIcon
    rm = (check_icon.value = vbChecked) And IconRemoveOn()

    key = IIf(rm, "1", "0") & "|" & mObjKey & "|" & Trim$(txt_icon_file.Text) & "|" & mSrcCaption
    If key = mPrevKey Then Exit Sub                 'nothing changed: keep the picture and the caption
    mPrevKey = key

    If mHasApp Then
        If mAppHas Then ok = DrawIconBytes(mIcoApp, mAppW, mAppH)
        If Not ok Then DrawPreview ""
        If rm Then
            lbl_icon_size.Caption = "Already replaced via APPICON. The VB6 icon will be removed"
        Else
            lbl_icon_size.Caption = "Already replaced via APPICON. The VB6 icon is still in the ressources"
        End If
    ElseIf rm Then
        If DrawSysIcon() Then
            lbl_icon_size.Caption = "Default EXE icon"
        Else
            DrawPreview ""
            lbl_icon_size.Caption = "Default EXE icon not available"
        End If
    ElseIf Len(Trim$(txt_icon_file.Text)) > 0 Then
        DrawPreview mSrcDraw
        lbl_icon_size.Caption = mSrcCaption
    Else
        If mObjHas Then ok = DrawIconBytes(mIcoObj, mObjW, mObjH)
        If ok Then
            lbl_icon_size.Caption = "VB6 icon in " & FileOf(MainObj())
        ElseIf Len(mObjKey) = 0 Then
            DrawPreview ""
            lbl_icon_size.Caption = "main.obj not found"
        ElseIf DrawSysIcon() Then
            lbl_icon_size.Caption = "Default EXE icon" & vbCrLf & "(no icon found in " & FileOf(MainObj()) & ")"
        Else
            DrawPreview ""
            lbl_icon_size.Caption = "No icon found in main.obj"
        End If
    End If

End Sub


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
    Dim pw As Long
    Dim ph As Long

    pw = pic_icon.ScaleWidth
    ph = pic_icon.ScaleHeight
    pic_icon.Line (0, 0)-(pw, ph), RGB(255, 255, 255), BF       'white base, only the gray squares are drawn on top
    For y = 0 To ph - 1 Step 8
        For x = IIf(((y \ 8) And 1) = 0, 8, 0) To pw - 1 Step 16
            pic_icon.Line (x, y)-(x + 7, y + 7), RGB(204, 204, 204), BF
        Next x
    Next y

End Sub

Private Sub DrawPreview(ByVal path As String)

    Dim img As Long
    Dim g As Long
    Dim w As Long
    Dim h As Long
    Dim dw As Long
    Dim dh As Long
    Dim ia As Long

    DrawChecker
    If mGdipToken <> 0 And Len(path) > 0 Then
        If GdipLoadImageFromFile(StrPtr(path), img) = 0 Then
            GdipGetImageWidth img, w
            GdipGetImageHeight img, h
            If w > 0 And h > 0 Then
                FitPreview w, h, dw, dh
                If GdipCreateFromHDC(pic_icon.hdc, g) = 0 Then
                    'enlarging: nearest neighbour keeps small icons crisp (bicubic smears them, DrawIconEx does not);
                    'shrinking: high quality bicubic
                    If dw > w Or dh > h Then GdipSetInterpolationMode g, 5 Else GdipSetInterpolationMode g, 7
                    GdipSetPixelOffsetMode g, 4                 'half pixel offset: no cut-off first row / column
                    'wrap mode TileFlipXY: without it GDI+ blends the outermost pixel row / column with
                    'transparency, which leaves a thin checkerboard line on the right and bottom edge
                    If GdipCreateImageAttributes(ia) = 0 Then GdipSetImageAttributesWrapMode ia, 3, 0, 0
                    GdipDrawImageRectRectI g, img, (pic_icon.ScaleWidth - dw) \ 2, (pic_icon.ScaleHeight - dh) \ 2, dw, dh, 0, 0, w, h, 2, ia, 0, 0
                    If ia <> 0 Then GdipDisposeImageAttributes ia: ia = 0
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
    mSrcDraw = ""
    mPrevKey = ""

    If Len(path) = 0 Then
        mSrcCaption = "No icon file selected"
    ElseIf Not FileExists(path) Then
        mSrcCaption = "File not found"
    ElseIf mGdipToken = 0 Then
        mSrcW = -1                                  'size unknown: every resolution may be chosen
        mSrcH = -1
        mSrcCaption = "GDI+ not available - no preview"
    ElseIf GdipSize(path, w, h) Then
        mSrcW = w
        mSrcH = h
        mSrcCaption = "Custom icon: " & w & " x " & h & " px"
        mSrcDraw = path
    Else
        mSrcCaption = "Cannot read this image"
    End If

    UpdateResEnable resetChecks
    UpdateIconPreview
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
    ScanObjIcon
    master = (check_icon.value = vbChecked) And Not IconRemoveOn() And Not mHasApp
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

'==============================================================================
' File version of a tool (version.dll)
'==============================================================================
Private Function HiW(ByVal v As Long) As Long
    If v < 0 Then
        HiW = ((v And &H7FFFFFFF) \ &H10000) Or &H8000&
    Else
        HiW = v \ &H10000
    End If
End Function

Private Function low(ByVal v As Long) As Long
    low = v And &HFFFF&
End Function

'"1.2.3.4" from the fixed file info of an exe / dll ("" when it has no version resource)
Private Function FileVersionOf(ByVal path As String) As String

    Dim dummy As Long
    Dim n As Long
    Dim buf() As Byte
    Dim pv As Long
    Dim ln As Long
    Dim v(0 To 3) As Long

    On Error GoTo EH
    If Not FileExists(path) Then Exit Function
    n = GetFileVersionInfoSize(path, dummy)
    If n <= 0 Then Exit Function
    ReDim buf(n - 1)
    If GetFileVersionInfo(path, 0&, n, buf(0)) = 0 Then Exit Function
    If VerQueryValue(buf(0), "\", pv, ln) = 0 Then Exit Function
    If pv = 0 Or ln < 16 Then Exit Function
    CopyMemory v(0), ByVal pv, 16
    If v(0) <> &HFEEF04BD Then Exit Function                 'VS_FIXEDFILEINFO signature
    FileVersionOf = HiW(v(2)) & "." & low(v(2)) & "." & HiW(v(3)) & "." & low(v(3))
EH:

End Function

'==============================================================================
' Project info: the VERSIONINFO resource (RT_VERSION) of main.obj
'==============================================================================
Private Function Al4(ByVal p As Long) As Long
    Al4 = ((p + 3) \ 4) * 4
End Function

'UTF-16 string of a byte array (at most maxChars characters, zero terminated)
Private Function WStrAt(a() As Byte, ByVal p As Long, ByVal maxChars As Long) As String

    Dim i As Long
    Dim c As Long
    Dim s As String

    For i = 0 To maxChars - 1
        If p + i * 2 + 1 > UBound(a) Then Exit For
        c = BU16(a, p + i * 2)
        If c = 0 Then Exit For
        s = s & ChrW$(c)
    Next i
    WStrAt = s

End Function

'First resource of a type of the loaded main.obj, whatever its id / language
Private Function ObjResFirst(ByVal resType As Long, data() As Byte) As Boolean

    Dim b1 As Long
    Dim s1 As Long
    Dim b2 As Long
    Dim s2 As Long
    Dim t As Long
    Dim d As Boolean
    Dim p As Long
    Dim off As Long
    Dim sz As Long
    Dim ix As Long
    Dim relOff As Long
    Dim relCnt As Long

    On Error GoTo EH
    If Not ObjSectionEx(".rsrc$01", b1, s1, ix, relOff, relCnt) Then Exit Function
    If Not ObjSection(".rsrc$02", b2, s2) Then Exit Function
    If Not ResDirFind(b1, 0, resType, "", t, d) Then Exit Function
    If Not d Then Exit Function
    Do While d                                      'type -> id -> language: first entry of every level
        p = b1 + t
        If U16(p + 12) + U16(p + 14) < 1 Then Exit Function
        d = (I32(p + 20) < 0)
        t = I32(p + 20) And &H7FFFFFFF
    Loop
    p = b1 + t
    off = ResDataOffset(b1, t, relOff, relCnt)      'offset inside .rsrc$02
    sz = I32(p + 4)
    If off < 0 Or sz <= 0 Or off + sz > s2 Then Exit Function
    ReDim data(sz - 1)
    CopyMemory data(0), mOb(b2 + off), sz
    ObjResFirst = True
EH:

End Function

'A string of the StringFileInfo block of a VS_VERSIONINFO resource ("" when missing)
Private Function VerStr(v() As Byte, ByVal wanted As String) As String

    Dim total As Long
    Dim p As Long
    Dim cl As Long
    Dim ck As String
    Dim t As Long
    Dim tl As Long
    Dim tk As String
    Dim s As Long
    Dim sl As Long
    Dim svl As Long
    Dim sk As String
    Dim vp As Long
    Dim r As String

    On Error GoTo EH
    total = BU16(v, 0)
    If total > UBound(v) + 1 Then total = UBound(v) + 1
    p = Al4(6 + (Len(WStrAt(v, 6, 40)) + 1) * 2)    'VS_FIXEDFILEINFO
    p = Al4(p + BU16(v, 2))                         'first child
    Do While p + 6 < total
        cl = BU16(v, p)
        If cl < 6 Then Exit Do
        ck = WStrAt(v, p + 6, 40)
        If ck = "StringFileInfo" Then
            t = Al4(p + 6 + (Len(ck) + 1) * 2)
            Do While t + 6 < p + cl
                tl = BU16(v, t)
                If tl < 6 Then Exit Do
                tk = WStrAt(v, t + 6, 20)
                s = Al4(t + 6 + (Len(tk) + 1) * 2)
                Do While s + 6 < t + tl
                    sl = BU16(v, s)
                    If sl < 6 Then Exit Do
                    svl = BU16(v, s + 2)
                    sk = WStrAt(v, s + 6, 64)
                    vp = Al4(s + 6 + (Len(sk) + 1) * 2)
                    If StrComp(sk, wanted, vbTextCompare) = 0 Then
                        r = Trim$(WStrAt(v, vp, svl))
                        If Len(r) > 0 Then VerStr = r: Exit Function
                    End If
                    s = Al4(s + sl)
                Loop
                t = Al4(t + tl)
            Loop
        End If
        p = Al4(p + cl)
    Loop
EH:

End Function

'"1.2.3.4" from the VS_FIXEDFILEINFO of a VS_VERSIONINFO resource
Private Function VerFixed(v() As Byte) As String

    Dim p As Long
    Dim sig As Long
    Dim ms As Long
    Dim ls As Long

    On Error GoTo EH
    p = Al4(6 + (Len(WStrAt(v, 6, 40)) + 1) * 2)
    If p + 16 > UBound(v) Then Exit Function
    CopyMemory sig, v(p), 4
    If sig <> &HFEEF04BD Then Exit Function
    CopyMemory ms, v(p + 8), 4
    CopyMemory ls, v(p + 12), 4
    VerFixed = HiW(ms) & "." & low(ms) & "." & HiW(ls) & "." & low(ls)
EH:

End Function

'Name / version / description / copyright of the project.
'Taken from the VERSIONINFO resource of main.obj (the VB6 "Make" tab); without it the name falls back
'to the output file (/OUT:) or to the file name of main.obj. The version is "" when main.obj has none.
Private Sub ProjectInfo(ByVal mainPath As String, ByRef pName As String, ByRef pVersion As String, ByRef pDesc As String, ByRef pCopy As String)

    Dim v() As Byte
    Dim nm As String

    pName = "": pVersion = "": pDesc = "": pCopy = ""

    If Len(mainPath) > 0 Then
        If ObjLoad(mainPath) Then
            If ObjResFirst(RT_VERSION, v) Then
                pName = VerStr(v, "ProductName")
                If Len(pName) = 0 Then pName = VerStr(v, "InternalName")
                pVersion = VerFixed(v)
                pDesc = VerStr(v, "FileDescription")
                If Len(pDesc) = 0 Then pDesc = VerStr(v, "Comments")
                pCopy = VerStr(v, "LegalCopyright")
            End If
        End If
    End If

    If Len(pName) = 0 Then
        nm = ArgValue("/OUT:")
        If Len(nm) > 0 Then pName = BaseOf(nm)
        If Len(pName) = 0 And Len(mainPath) > 0 Then pName = BaseOf(mainPath)
    End If

End Sub

'==============================================================================
' Manifest (VBMF)
'==============================================================================
Private Function MainObj() As String
    MainObj = StripQ(txt_main_obj.Text)
End Function

Private Function ComboText(cbo As ComboBox) As String
    If cbo.ListIndex >= 0 Then ComboText = cbo.List(cbo.ListIndex) Else ComboText = "default"
End Function

Private Sub SelectCombo(cbo As ComboBox, ByVal Text As String)

    Dim i As Long

    cbo.ListIndex = 0
    For i = 0 To cbo.ListCount - 1
        If StrComp(cbo.List(i), Text, vbTextCompare) = 0 Then cbo.ListIndex = i: Exit For
    Next i

End Sub

'Text safe for XML (pure ASCII: everything else becomes a character reference)
Private Function XmlEsc(ByVal s As String) As String

    Dim i As Long
    Dim c As Long
    Dim r As String

    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case c
            Case 38: r = r & "&amp;"
            Case 60: r = r & "&lt;"
            Case 62: r = r & "&gt;"
            Case 34: r = r & "&quot;"
            Case 39: r = r & "&apos;"
            Case 0 To 31: r = r & " "
            Case 127 To 65535: r = r & "&#" & c & ";"
            Case Else: r = r & Chr$(c)
        End Select
    Next i
    XmlEsc = r

End Function

'assemblyIdentity name: letters, digits and . - _ only ("My App" -> "My.App")
Private Function ManifestName(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim r As String

    s = Trim$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "." Or ch = "-" Or ch = "_" Then
            r = r & ch
        ElseIf ch = " " Then
            r = r & "."
        Else
            r = r & "_"
        End If
    Next i
    Do While Len(r) > 0                                      'no leading / trailing "." or "_"
        If Right$(r, 1) = "." Or Right$(r, 1) = "_" Then r = Left$(r, Len(r) - 1) Else Exit Do
    Loop
    Do While Len(r) > 0
        If Left$(r, 1) = "." Or Left$(r, 1) = "_" Then r = Mid$(r, 2) Else Exit Do
    Loop
    If Len(r) = 0 Then r = "App"
    ManifestName = r

End Function

'Name / version / description that go into the manifest (filled in from the project info)
Private Sub ManifestIdent(ByRef nm As String, ByRef ver As String, ByRef ds As String)

    Dim cp As String

    ProjectInfo MainObj(), nm, ver, ds, cp
    If Len(ver) = 0 Then ver = "1.0.0.0"
    nm = ManifestName(nm)

End Sub

'Supported OS GUIDs (1 = Vista ... 5 = Windows 10 / 11)
Private Function OsGuid(ByVal i As Long) As String
    Select Case i
        Case 1: OsGuid = "{e2011457-1546-43c5-a5fe-008deee3d3f0}"
        Case 2: OsGuid = "{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"
        Case 3: OsGuid = "{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"
        Case 4: OsGuid = "{1f676c76-80e1-4239-95bb-83d0f6d0da78}"
        Case Else: OsGuid = "{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"
    End Select
End Function

Private Function BuildManifest() As String

    Dim nm As String
    Dim ver As String
    Dim ds As String
    Dim s As String
    Dim ws As String
    Dim i As Long

    ManifestIdent nm, ver, ds

    s = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>" & vbCrLf
    s = s & "<assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'>" & vbCrLf
    s = s & "  <assemblyIdentity type='win32' name='%NAME%' version='%VER%' processorArchitecture='x86'/>" & vbCrLf
    If Len(ds) > 0 Then s = s & "  <description>%DESC%</description>" & vbCrLf

    If check_vstyles.value = vbChecked Then                  'Common Controls 6 = visual styles
        s = s & "  <dependency>" & vbCrLf
        s = s & "    <dependentAssembly>" & vbCrLf
        s = s & "      <assemblyIdentity type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'/>" & vbCrLf
        s = s & "    </dependentAssembly>" & vbCrLf
        s = s & "  </dependency>" & vbCrLf
    End If

    If cbo_uac.ListIndex > 0 Then
        s = s & "  <trustInfo xmlns='urn:schemas-microsoft-com:asm.v3'>" & vbCrLf
        s = s & "    <security>" & vbCrLf
        s = s & "      <requestedPrivileges>" & vbCrLf
        s = s & "        <requestedExecutionLevel level='" & ComboText(cbo_uac) & "' uiAccess='false'/>" & vbCrLf
        s = s & "      </requestedPrivileges>" & vbCrLf
        s = s & "    </security>" & vbCrLf
        s = s & "  </trustInfo>" & vbCrLf
    End If

    If cbo_os.ListIndex > 0 Then                             'every OS up to the selected one
        s = s & "  <compatibility xmlns='urn:schemas-microsoft-com:compatibility.v1'>" & vbCrLf
        s = s & "    <application>" & vbCrLf
        For i = 1 To cbo_os.ListIndex
            s = s & "      <supportedOS Id='" & OsGuid(i) & "'/>" & vbCrLf
        Next i
        s = s & "    </application>" & vbCrLf
        s = s & "  </compatibility>" & vbCrLf
    End If

    If check_dpi.value = vbChecked Then
        ws = ws & "      <dpiAware xmlns='http://schemas.microsoft.com/SMI/2005/WindowsSettings'>true</dpiAware>" & vbCrLf
    End If
    If check_segheap.value = vbChecked Then
        ws = ws & "      <heapType xmlns='http://schemas.microsoft.com/SMI/2020/WindowsSettings'>SegmentHeap</heapType>" & vbCrLf
    End If
    If Len(ws) > 0 Then
        s = s & "  <application xmlns='urn:schemas-microsoft-com:asm.v3'>" & vbCrLf
        s = s & "    <windowsSettings>" & vbCrLf
        s = s & ws
        s = s & "    </windowsSettings>" & vbCrLf
        s = s & "  </application>" & vbCrLf
    End If
    s = s & "</assembly>" & vbCrLf

    s = Replace(s, "'", Chr$(34))                            'the template uses ' instead of "
    s = Replace(s, "%NAME%", nm)
    s = Replace(s, "%VER%", ver)
    s = Replace(s, "%DESC%", XmlEsc(ds))
    BuildManifest = s

End Function

Private Function DefaultManifestPath() As String

    Dim m As String
    Dim d As String
    Dim o As String
    Dim nm As String

    m = MainObj()
    o = ArgValue("/OUT:")
    If Len(m) > 0 Then d = DirOf(m)
    If Len(d) = 0 Then d = mCfgDir

    If Len(o) > 0 Then
        nm = FileOf(o)                                       'app.exe -> app.exe.manifest
    ElseIf Len(m) > 0 Then
        If IsDllBuild() Then nm = BaseOf(m) & ".dll" Else nm = BaseOf(m) & ".exe"
    Else
        nm = "app.exe"
    End If
    DefaultManifestPath = AddBS(d) & nm & ".manifest"

End Function

Private Sub btn_manifest_browse_Click()

    Dim c As Collection

    Set c = FileDialog("Select the manifest file", "Manifest files (*.manifest)|*.manifest|All files (*.*)|*.*", False, True, txt_manifest.Text)
    If c Is Nothing Then Exit Sub
    txt_manifest.Text = c(1)

End Sub

Private Sub btn_createmanifest_Click()

    Dim nm As String
    Dim ver As String
    Dim ds As String

    If CreateManifest() Then
        ManifestIdent nm, ver, ds
        MsgBox "Manifest created:" & vbCrLf & Trim$(txt_manifest.Text) & vbCrLf & vbCrLf & _
               "Name: " & nm & vbCrLf & "Version: " & ver, vbInformation
    End If

End Sub

'Writes the manifest file (name / version / description come from the project info)
Private Function CreateManifest() As Boolean

    Dim p As String
    Dim f As Long
    Dim msg As String

    p = Trim$(txt_manifest.Text)
    If Len(p) = 0 Then
        p = DefaultManifestPath()
        txt_manifest.Text = p
    End If

    On Error GoTo EH
    f = FreeFile
    Open p For Output As #f
    Print #f, BuildManifest();
    Close #f
    CreateManifest = True
    Exit Function

EH:
    msg = Err.Description                            'read first: any On Error statement clears Err
    On Error Resume Next
    Close #f
    MsgBox "Cannot write " & p & vbCrLf & msg, vbExclamation

End Function

'Case-insensitive ASCII search in the loaded object (low = lower-case needle)
Private Function ObjHasText(ByVal low As String) As Boolean

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim m As Long
    Dim f As Long
    Dim c As Long

    m = Len(low)
    If m = 0 Then Exit Function
    f = Asc(low)
    n = UBound(mOb) + 1 - m
    For i = 0 To n
        c = mOb(i)
        If c >= 65 And c <= 90 Then c = c + 32
        If c = f Then
            For j = 1 To m - 1
                c = mOb(i + j)
                If c >= 65 And c <= 90 Then c = c + 32
                If c <> Asc(Mid$(low, j + 1, 1)) Then Exit For
            Next j
            If j >= m Then ObjHasText = True: Exit Function
        End If
    Next i

End Function

Private Function ManifestSummary() As String

    Dim s As String

    If check_vstyles.value = vbChecked Then s = s & ", visual styles"
    If check_dpi.value = vbChecked Then s = s & ", DPI aware"
    If check_segheap.value = vbChecked Then s = s & ", segment heap"
    If cbo_uac.ListIndex > 0 Then s = s & ", UAC " & ComboText(cbo_uac)
    If cbo_os.ListIndex > 0 Then s = s & ", " & ComboText(cbo_os)
    If Len(s) > 0 Then s = Mid$(s, 3) Else s = "default settings"
    ManifestSummary = s

End Function

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

    Dim v As String

    If FileExists(path) Then
        v = FileVersionOf(path)
        If Len(v) > 0 Then v = "v" & v Else v = "no version"
        ToolState = "found    " & Left$(v & Space$(14), 14) & "  " & path
    Else
        ToolState = "MISSING  " & path
    End If

End Function

Private Function OnOff(ByVal b As Boolean) As String
    If b Then OnOff = "on" Else OnOff = "off"
End Function

'value of the first checked argument of a tool list that starts with prefix (e.g. "/OUT:")
Private Function ToolArgValue(ByVal tool As Long, ByVal prefix As String) As String

    Dim i As Long
    Dim t As String

    For i = 0 To args_list(tool).ListCount - 1
        If args_list(tool).Selected(i) Then
            t = NormSw(args_list(tool).List(i))
            If UCase$(Left$(t, Len(prefix))) = UCase$(prefix) Then
                ToolArgValue = StripQ(Mid$(t, Len(prefix) + 1))
                Exit Function
            End If
        End If
    Next i

End Function

Private Function ArgValue(ByVal prefix As String) As String
    ArgValue = ToolArgValue(TOOL_LINK, prefix)
End Function

'Is this exact switch among the checked arguments of a tool list?
Private Function ToolHasArg(ByVal tool As Long, ByVal arg As String) As Boolean

    Dim i As Long

    For i = 0 To args_list(tool).ListCount - 1
        If args_list(tool).Selected(i) Then
            If UCase$(StripQ(NormSw(args_list(tool).List(i)))) = UCase$(arg) Then ToolHasArg = True: Exit Function
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
    Dim ic As String
    Dim isDll As Boolean
    Dim pName As String
    Dim pVer As String
    Dim pDesc As String
    Dim pCopy As String

    isDll = IsDllBuild()
    If isDll Then
        If opt_dll(1).value Then kind = "DLL (Standard)" Else kind = "DLL (ActiveX)"
    Else
        kind = "EXE"
    End If
    main = MainObj()
    If Len(main) = 0 Then main = FindMainObj()
    For i = 0 To args_list(TOOL_LINK).ListCount - 1
        If args_list(TOOL_LINK).Selected(i) Then nOn = nOn + 1
    Next i

    ProjectInfo main, pName, pVer, pDesc, pCopy

    s = "PROJECT" & vbCrLf
    s = s & InfoLine("Name", pName)
    s = s & InfoLine("Version", pVer)
    s = s & InfoLine("Description", pDesc)
    s = s & InfoLine("Copyright", pCopy)
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
    If Len(main) = 0 Then
        s = s & InfoLine("main.obj", "auto detection failed")
    ElseIf FileExists(main) Then
        s = s & InfoLine("main.obj", main)
        s = s & InfoLine("main.obj size", Format$(FileLen(main), "#,##0") & " bytes")
    Else
        s = s & InfoLine("main.obj", "not found: " & main)
    End If

    ScanObjIcon
    If mHasApp Then
        ic = "APPICON"
        If mAppHas Then ic = ic & ", " & IconInfo(mAppW, mAppH, mAppBits)
    ElseIf mObjHas Then
        ic = "VB6 icon, " & IconInfo(mObjW, mObjH, mObjBits)
    Else
        ic = "none"
    End If
    s = s & InfoLine("Icon", ic)
    s = s & vbCrLf & "BUILD STEPS" & vbCrLf
    s = s & InfoLine("1. VBExtern", OnOff(check_extern.value = vbChecked))
    If IconStepOn() Then
        If EffectiveRemove() Then
            s = s & InfoLine("2. VBIcon", "on - remove the VB6 icon only")
        Else
            s = s & InfoLine("2. VBIcon", "on - " & FileOf(Trim$(txt_icon_file.Text)) & "  sizes " & ResList() & "  depth " & DepthList())
        End If
    Else
        s = s & InfoLine("2. VBIcon", "off")
    End If
    If check_static.value = vbChecked Then
        s = s & InfoLine("3. VBStatic", "on - runtime " & mRuntime & IIf(check_noicon.value = vbChecked, "  (/NOICON)", ""))
    Else
        s = s & InfoLine("3. VBStatic", "off")
    End If
    If check_manifest.value = vbChecked Then
        s = s & InfoLine("4. VBMF", "on - " & FileOf(Trim$(txt_manifest.Text)) & "  (" & ManifestSummary() & ")")
    Else
        s = s & InfoLine("4. VBMF", "off")
    End If
    s = s & InfoLine("5. LINK2", "always")

    s = s & vbCrLf & "TOOLS" & vbCrLf
    s = s & InfoLine("LINK2.EXE", ToolState(mToolLink))
    s = s & InfoLine("VBExtern.exe", ToolState(mToolExtern))
    s = s & InfoLine("VBIcon.exe", ToolState(mToolIcon))
    s = s & InfoLine("VBStatic.exe", ToolState(mToolStatic))
    s = s & InfoLine("VBMF.exe", ToolState(mToolManifest))
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
    usesRt = ObjHasText("msvbvm60")
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

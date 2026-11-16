!include "MUI2.nsh"

OutFile "LinkAlechemyVB6_Setup_win32.exe"
Name "LinkAlchemy"
BrandingText "LinkAlchemy for VB6 Setup"
RequestExecutionLevel admin

;info
VIProductVersion "1.0.0.0"
VIAddVersionKey "ProductName" "LinkAlchemy"
VIAddVersionKey "LegalCopyright" "Copyright © 2026"
VIAddVersionKey "FileDescription" "LinkAlchemy for VB6 - Setup"
VIAddVersionKey "FileVersion" "1.0.0"

;style
!define MUI_ICON "icon.ico"
!define MUI_WELCOMEFINISHPAGE_BITMAP "welcome.bmp"
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_RIGHT
!define MUI_HEADERIMAGE_BITMAP "banner.bmp"
!define MUI_WELCOMEPAGE_TEXT "This wizard will guide you through the installation of$\r$\nLinkAlchemy.$\r$\n$\r$\nIt is recommended that you close all other applications$\r$\nbefore starting Setup. This includes mainly VB6 and$\r$\nLINK.EXE.$\r$\n$\r$\nClick Next to continue."
!define MUI_DIRECTORYPAGE_TEXT_TOP "Please select the installation directory where LinkAlchemy and its associated files and components should be installed. Ensure that the selected directory corresponds to your existing VB6 installation."
!define MUI_DIRECTORYPAGE_TEXT_DESTINATION "Select a VB6 installation directory:"
!define MUI_DIRECTORYPAGE_VERIFYONLEAVE
!define MUI_FINISHPAGE_TEXT "LinkAlchemy has been sucessfully installed to VB6."
ShowInstDetails show

Function .onInit

	;auto detect vb6 path
	SetRegView 32
	ReadRegStr $INSTDIR HKLM "SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic" "ProductDir"

FunctionEnd



; forms
!insertmacro MUI_PAGE_WELCOME
!define MUI_PAGE_CUSTOMFUNCTION_LEAVE PreInstallCheck
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

Function PreInstallCheck
    IfFileExists "$INSTDIR\LINK.EXE" 0 MissingFiles
    IfFileExists "$INSTDIR\VB6.EXE" 0 MissingFiles

    Push "VB6.EXE"
    Call IsProcessRunning
    Pop $0
    StrCmp $0 1 VB6Running

    Push "LINK.EXE"
    Call IsProcessRunning
    Pop $0
    StrCmp $0 1 LINKRunning

    Return

    MissingFiles:
        MessageBox MB_ICONSTOP "Not a VB6 installation directory!"
        Abort

    VB6Running:
        MessageBox MB_ICONSTOP "VB6.EXE is running. Save your Projects and close VB6 before installing."
        Abort

    LINKRunning:
        MessageBox MB_ICONSTOP "LINK.EXE is running. Close it before installing."
        Abort
FunctionEnd




Section InstallFiles
	SetOutPath $INSTDIR
    	SetOverwrite ifnewer

	File "..\vbextern\VBExtern.exe"
	File "..\vbicon\VBIcon.exe"
	File "..\vbmf\VBMF.exe"
	File "..\vbstatic\VBStatic.exe"
	
	IfFileExists "$INSTDIR\LINK2.EXE" SkipRename
	Rename "$INSTDIR\LINK.EXE" "$INSTDIR\LINK2.EXE"
SkipRename:
	File "..\link\LINK.EXE"

SectionEnd

















Function IsProcessRunning
    Exch $R0                ; $R0 = process name
    Push $R1
    Push $R2
    Push $R3
    Push $R4
    Push $R5
    StrCpy $R4 0            ; result

    ; TH32CS_SNAPPROCESS = 2
    System::Call 'kernel32::CreateToolhelp32Snapshot(i 2, i 0) i .R1'
    ; PROCESSENTRY32 (ANSI, x86) = 296 bytes
    System::Call '*(i 296, i 0, i 0, i 0, i 0, i 0, i 0, i 0, i 0, &t260 "") i .R2'
    System::Call 'kernel32::Process32First(i R1, i R2) i .R3'

    loop:
        StrCmp $R3 0 done
        System::Call '*$R2(i, i, i, i, i, i, i, i, i, &t260 .R5)'
        StrCmp $R5 $R0 found        ; StrCmp is case-insensitive
        System::Call 'kernel32::Process32Next(i R1, i R2) i .R3'
        Goto loop

    found:
        StrCpy $R4 1

    done:
        System::Call 'kernel32::CloseHandle(i R1)'
        System::Free $R2

        StrCpy $R0 $R4
        Pop $R5
        Pop $R4
        Pop $R3
        Pop $R2
        Pop $R1
        Exch $R0                    ; restore old $R0, leave result on stack
FunctionEnd








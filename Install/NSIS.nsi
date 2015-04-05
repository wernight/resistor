;!include UpgradeDLL.nsi
;!include AddSharedDLL.nsi
;!include un.RemoveSharedDLL.nsi
!include "Library.nsh"

!define PRODUCT_NAME "Résistor"	;Define your own software name here
!define VERSION "2.31"		;Define your own software version here

;--------------------------------
;Configuration

	Name ${PRODUCT_NAME}

	;Do A CRC Check
	CRCCheck On

	;Compression format
	SetCompressor /SOLID lzma

	;Output File Name
	OutFile "resistor-v${VERSION}-setup.exe"

	;The Default Installation Directory
	InstallDir "$PROGRAMFILES\Résistor"

	;Remember install folder
	InstallDirRegKey HKCU "Software\ALC-WBC\${PRODUCT_NAME}" ""

;--------------------------------
;Interface Settings

LicenseData "Licence.txt"

;--------------------------------
;Pages

Page license
Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------
;Languages
 
LoadLanguageFile "${NSISDIR}\Contrib\Language files\French.nlf"

;--------------------------------
;Installer Sections

Section "Résistor (requis)"
	SectionIn RO

	;Install Files
	SetOutPath $INSTDIR
	File "..\Resistor.exe"
	File "Resistor.cle"

	; Write the uninstall keys for Windows
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "DisplayName" "${PRODUCT_NAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "DisplayIcon" "$INSTDIR\Resistor.exe"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "DisplayVersion" "${VERSION}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "Publisher" "Beroux"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "URLInfoAbout" "http://www.beroux.com/"	;Publisher's link
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "HelpLink" "http://www.beroux.com/france/logiciels/resistor/" ;Support Information
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "URLUpdateInfo" "http://www.beroux.com/france/logiciels/resistor/"	;Product Updates
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "NoModify" 0x00000001
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" "NoRepair" 0x00000001

	;Store installation folder
	WriteRegStr HKCU "Software\ALC-WBC\${PRODUCT_NAME}" "" $INSTDIR

	;Create uninstaller
	WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Var ALREADY_INSTALLED

Section "-Install VB5 runtimes"
	;Add code here that sets $ALREADY_INSTALLED to a non-zero value if the application is already installed. For example:

	IfFileExists "$INSTDIR\Resistor.exe" 0 new_installation ;Replace Resistor.exe with your application filename
		StrCpy $ALREADY_INSTALLED 1
	new_installation:

	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_NOTPROTECTED "DLL\msvbvm50.dll" "$SYSDIR\msvbvm50.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\oleaut32.dll" "$SYSDIR\oleaut32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\olepro32.dll" "$SYSDIR\olepro32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\comcat.dll"   "$SYSDIR\comcat.dll"   "$SYSDIR"
	!insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\asycfilt.dll" "$SYSDIR\asycfilt.dll" "$SYSDIR"
	!insertmacro InstallLib TLB    $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\stdole2.tlb"  "$SYSDIR\stdole2.tlb"  "$SYSDIR"

	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\VB5FR.dll"    "$SYSDIR\VB5FR.dll"    "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\COMDLG32.OCX" "$SYSDIR\COMDLG32.OCX" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\THREED32.OCX" "$SYSDIR\THREED32.OCX" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\CmDlgFR.dll"  "$SYSDIR\CmDlgFR.dll"  "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\PcClpFR.dll"  "$SYSDIR\PcClpFR.dll"  "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\PICCLP32.OCX" "$SYSDIR\PICCLP32.OCX" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\TabCtFR.dll"  "$SYSDIR\TabCtFR.dll"  "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\TABCTL32.ocx" "$SYSDIR\TABCTL32.ocx" "$SYSDIR"

	; Shared WBSCrypte DLL
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED    "DLL\WBCCrypteDLL.dll" "$SYSDIR\WBCCrypteDLL.dll" "$SYSDIR"
SectionEnd

Section "-un.Uninstall VB5 runtimes"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\msvbvm50.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\oleaut32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\olepro32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comcat.dll"
	!insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\asycfilt.dll"
	!insertmacro UnInstallLib TLB    SHARED NOREMOVE "$SYSDIR\stdole2.tlb"

	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\VB5FR.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\COMDLG32.OCX"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\THREED32.OCX"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\CmDlgFR.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\PcClpFR.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\PICCLP32.OCX"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\TabCtFR.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\TABCTL32.ocx"

	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\WBCCrypteDLL.dll"
SectionEnd

Section "Raccourcis"
	;Add Shortcuts
	CreateDirectory "$SMPROGRAMS\${PRODUCT_NAME}"
	CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Résistor.lnk" "$INSTDIR\Resistor.exe" "" "$INSTDIR\Resistor.exe" 0
	WriteINIStr "$SMPROGRAMS\${PRODUCT_NAME}\Site web de Résistor.url" "InternetShortcut" "URL" "http://www.beroux.com/france/logiciels/resistor/"
	CreateShortCut "$SMPROGRAMS\${PRODUCT_NAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe" "" "$INSTDIR\Uninstall.exe" 0
SectionEnd

Section Uninstall
	;Delete Files
	Delete "$INSTDIR\Resistor.exe"
	Delete "$INSTDIR\Resistor.cle"
	Delete "$INSTDIR\Resistor.dat"

	;Delete Start Menu Shortcuts
	Delete "$SMPROGRAMS\Résistor\*.*"
	RmDir "$SMPROGRAMS\Résistor"

	;Delete Uninstaller And Unistall Registry Entries
	Delete "$INSTDIR\Uninstall.exe"
	RMDir "$INSTDIR"

	DeleteRegKey HKLM "SOFTWARE\Résistor"
	DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
	DeleteRegKey /ifempty HKCU "Software\ALC-WBC\${PRODUCT_NAME}"
	DeleteRegKey /ifempty HKCU "Software\ALC-WBC"
SectionEnd

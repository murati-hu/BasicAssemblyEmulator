!include "MUI.nsh"

;--------------------------------
;Configuration

Name "Basic Assembly Emulator 8 - Mur�ti �kos"  
OutFile "base8_setup.exe"

  ShowInstDetails show

  InstallDir "$PROGRAMFILES\BAsE8"
  
  InstallDirRegKey HKCU "Software\BAsE8" ""

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "Hungarian"
  
;--------------------------------
;Language Strings

  ;Description
	LangString DESC_base8 ${LANG_HUNGARIAN} "Basic Assembly Emulator v0.1"
	LangString DESC_nyelvek ${LANG_HUNGARIAN} "Telep�thet� nyelvek: Angol"
	LangString DESC_peldak ${LANG_HUNGARIAN} "Egyszer� assembly p�ldaprogramok"
	LangString DESC_VB6 ${LANG_HUNGARIAN} "Fut�shoz sz�ks�ges Visual Basic 6.0 (SP5) Runtime f�jlok telep�t�se.(Win XP alatt nem sz�ks�ges)"
	LangString DESC_Eltavolit ${LANG_HUNGARIAN} "Elt�vol�t� alkalmaz�s telep�t�se. (Uninstall)"

;--------------------------------
;Installer Sections

Section "BAsE8 v0.1" base8
	SectionIn RO

	detailprint ">>> Microsoft Common dialog ActiveX vez�rl� telep�t�se..."
	setoutpath $SYSDIR
	file "comdlg32.ocx"
	execwait "regsvr32.exe /i /s $SYSDIR/comdlg32.ocx"
	detailprint ""
	

	detailprint ">>> Program telep�t�se..."
  	SetOutPath "$INSTDIR"
	File "..\base8.exe"
  	CreateDirectory "$SMPROGRAMS\BAsE8"
	CreateShortCut "$SMPROGRAMS\BAsE8\Basic Assembly Emulator 8.lnk" "$INSTDIR\base8.exe"

	detailprint ""
SectionEnd

Section "T�bbnyelv�s�g" nyelvek
	detailprint ">>> Idegen nyelvek m�sol�sa..."
	createdirectory "$INSTDIR\nyelvek"
	SetOutPath "$INSTDIR\nyelvek"
	file "..\nyelvek\*.*"
	detailprint ""
sectionend

Section "P�ldaprogramok" peldak
	detailprint ">>> P�ldaprogramok m�sol�sa..."
	createdirectory "$INSTDIR\peldak"
	SetOutPath "$INSTDIR\peldak"
	file "..\asm\*.*"
	detailprint ""
sectionend

section "Microsoft Visual Basic 6.0 Runtime (SP5)" VB6
	detailprint ">>> Microsoft Visual Basic 6.0 Runtime (SP5) telep�t�se..."
	setoutpath $SYSDIR
	file "vbrun.exe"
	execwait "$SYSDIR\vbrun.exe /q"
	detailprint ""
sectionend

Section "Elt�vol�t� alkalmaz�s" Eltavolit
	detailprint ">>> Elt�vo�t� alkalmaz�s telep�t�se..."
	SetOutPath "$INSTDIR"
	WriteUninstaller "$INSTDIR\eltavolit.exe"
	CreateShortCut "$SMPROGRAMS\BAsE8\Elt�vol�t�s.lnk" "$INSTDIR\eltavolit.exe" 
Sectionend 


;!insertmacro MUI_SECTIONS_FINISHHEADER


!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
	!insertmacro MUI_DESCRIPTION_TEXT ${base8} $(DESC_base8)
	!insertmacro MUI_DESCRIPTION_TEXT ${nyelvek} $(DESC_nyelvek)
	!insertmacro MUI_DESCRIPTION_TEXT ${peldak} $(DESC_peldak)
	!insertmacro MUI_DESCRIPTION_TEXT ${VB6} $(DESC_VB6)
	!insertmacro MUI_DESCRIPTION_TEXT ${Eltavolit} $(DESC_Eltavolit)
!insertmacro MUI_FUNCTION_DESCRIPTION_END
 
;--------------------------------
;Uninstaller Section

Section "Uninstall"
	delete "$INSTDIR\*.*"
	delete "$INSTDIR\nyelvek\*.*"
	delete "$INSTDIR\peldak\*.*"
	delete "$SMPROGRAMS\BAsE8\*.*"
	rmdir "$SMPROGRAMS\BAsE8"
	rmdir "$INSTDIR\nyelvek"
	rmdir "$INSTDIR\peldak"
	rmdir "$INSTDIR"
  	;!insertmacro MUI_UNFINISHHEADER
SectionEnd
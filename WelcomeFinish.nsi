;NSIS Modern User Interface
;Welcome/Finish Page Example Script
;Written by Joost Verburg

;--------------------------------
;Include Modern UI

  !include "MUI2.nsh"

;--------------------------------
;General

  ;Name and file
  Name "Infinity"
  OutFile "Infinity_Installer_EN.exe"

  ;Default installation folder
  InstallDir "$%systemdrive%\Program files\Infinity"
  
  ;Get installation folder from registry if available
  InstallDirRegKey HKCU "Software\Infinity" ""

  ;Request application privileges for Windows Vista
  RequestExecutionLevel user

;--------------------------------
;Variables

  Var StartMenuFolder
;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_WELCOME
  !insertmacro MUI_PAGE_LICENSE "${NSISDIR}\Docs\Modern UI\License.txt"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY

  ;Start Menu Folder Page Configuration
  !define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU" 
  !define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Infinity" 
  !define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"

  !insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder

  !insertmacro MUI_PAGE_INSTFILES
  !insertmacro MUI_PAGE_FINISH

  !insertmacro MUI_UNPAGE_WELCOME
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  !insertmacro MUI_UNPAGE_FINISH

;--------------------------------
;Languages

  !insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections

Section "Server Files" SecServer
  
  SetOutPath "$INSTDIR\Server\Database" 
  file Database\Data.mdb

  ;ADD YOUR OWN FILES HERE...
  
  SetOutPath "$INSTDIR\Server" 
  file Infinity.exe

  ;Store installation folder
  WriteRegStr HKCU "Software\Infinity\Server" "" $INSTDIR

  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Server\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "DisplayIcon" "C:\Program Files\Infinity\Server\Infinity.exe,0"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "DisplayName" "Infinity" 
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "DisplayVersion" "1.0.0" 
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "InstallLocation" "C:\Program Files\Infinity"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "UninstallString" "C:\Program Files\Infinity\Server\Uninstall.exe"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "NoModify" "1"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity" "NoRepair" "1"

  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
    
    ;Create shortcuts
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder\Server"
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Server\Infinity.lnk" "$INSTDIR\Server\Infinity.exe"
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Server\Uninstall.exe"
  
  !insertmacro MUI_STARTMENU_WRITE_END

SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_SecServer ${LANG_ENGLISH} "Installs Servers Files."

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecServer} $(DESC_SecServer)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

Section "Uninstall"

  ;ADD YOUR OWN FILES HERE...
  Delete "$INSTDIR\Database\Data.mdb"   
  Delete "$INSTDIR\Infinity.exe"
  Delete "$INSTDIR\Uninstall.exe"

  RMDir "$INSTDIR\Database"
  RMDir "$INSTDIR\Server"
  RMDir "$INSTDIR"

 !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder
  
  Delete "$SMPROGRAMS\$StartMenuFolder\Server\Infinity.lnk"
  Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk"
  RMDir "$SMPROGRAMS\$StartMenuFolder\Server"
  RMDir "$SMPROGRAMS\$StartMenuFolder"

  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Infinity"
  DeleteRegKey  HKCU "Software\infinity\Server"
  DeleteRegKey /ifempty HKCU "Software\infinity"

SectionEnd

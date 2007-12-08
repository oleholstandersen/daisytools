!define VBFILESDIR "..\bin\vbruntimes"

!include "extras.nsh"
!include "vbsections.nsh"
!include "xmlcheck.nsh"

!Define ProgName "Daisy 2.02 Validator"
!Define ProgNameUiIterator "Daisy 2.02 Iterator Validator"
!Define ProgVer "20050401"

OutFile "Daisy.2.02.Validator.${ProgVer}.Install.exe"
Name "${ProgName} ${ProgVer} Installer"

LicenseData "..\bin\ui\lgpl.txt"
LicenseText "The ${ProgName} is under the following license, please press \
  'i agree' to agree to the licensetext." "I agree"

ShowInstDetails show

!Define BinDirUi "..\bin\ui\"
!Define BinDirDll "..\bin\dll\"
!Define ExtFilDir "..\bin\ui\externals"
!Define SysFilDir "..\bin\sysfiles\"
!Define ManFilDir "..\bin\ui\manual\"

ComponentText "Please select which components you want installed."
InstallDir "$PROGRAMFILES\DaisyWare\${ProgName}\"
DirText "Select the directory to install ${ProgName} into"

Function .onInit
LookForNSIS:
  ReadRegStr $R1 HKLM \
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}\" \
    "UninstallString"

  strCmp $R1 "" PrevNSISNotFound

  MessageBox MB_YESNO "A version of ${ProgName} is already installed. It is STRONGLY recommended to uninstall the previous version first. Quit install?" IDNO PrevISNotFound
  Quit

PrevNSISNotFound:
  ReadRegStr $R0 HKLM \
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{932E1C7A-2EF9-4F84-BFAC-6364C2D6BC48}" \
    "UninstallString"
  strCmp $R0 "" PrevISNotFound

  MessageBox MB_YESNO "A version of ${ProgName} is already installed. It is STRONGLY recommended to uninstall the previous version first. Quit install?" IDNO PrevISNotFound
  Quit
PrevISNotFound:
  call IsMSXML40Installed
FunctionEnd

Page license
Page components
page directory
page instfiles

Section "Core files"
  SectionIn RO  ; RO means Read Only

  setOverwrite ifnewer

  DeleteRegValue HKLM "SOFTWARE\DaisyWare\validator\Misc" "AppPath"
  DeleteRegValue HKLM "SOFTWARE\DaisyWare\validator\Misc" "DefRepPath"
  DeleteRegValue HKLM "SOFTWARE\DaisyWare\validator\Misc" "Dtd_AdtdPath"
  DeleteRegValue HKLM "SOFTWARE\DaisyWare\validator\Misc" "TempPath"
  
  DeleteRegValue HKCU "SOFTWARE\DaisyWare\validator\Misc" "AppPath"
  DeleteRegValue HKCU "SOFTWARE\DaisyWare\validator\Misc" "DefRepPath"
  DeleteRegValue HKCU "SOFTWARE\DaisyWare\validator\Misc" "Dtd_AdtdPath"
  DeleteRegValue HKCU "SOFTWARE\DaisyWare\validator\Misc" "TempPath"

  SetOutPath "$SYSDIR"

  !define UPGRADEDLL_NOREGISTER
  !insertmacro UpgradeDll "${SysFilDir}\comdlg32.ocx" "$SYSDIR\comdlg32.ocx" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${SysFilDir}\richtx32.ocx" "$SYSDIR\richtx32.ocx" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${SysFilDir}\mscomctl.ocx" "$SYSDIR\mscomctl.ocx" "$SYSDIR\temp"

  call InstallVb6Runtimes

  SetOutPath "$INSTDIR" ; set target path for file commands below
  File "${BinDirUi}*.*" ; compile this to installer exe, and extract when installing

  CreateDirectory "$SYSDIR\DaisyWare"
  !insertmacro UpgradeDll "${BinDirDll}validatorengine.dll" \
    "$SYSDIR\DaisyWare\validatorengine.dll" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${BinDirDll}dtdparser.dll" \
     "$SYSDIR\DaisyWare\dtdparser.dll" "$SYSDIR\temp"

  regDll "$SYSDIR\DaisyWare\validatorengine.dll"
  regDll "$SYSDIR\DaisyWare\dtdparser.dll"

  Push "$SYSDIR\DaisyWare\validatorengine.dll"
  Call AddSharedDll
  Push "$SYSDIR\DaisyWare\dtdparser.dll"
  Call AddSharedDll

  WriteUninstaller "$INSTDIR\Uninstall.exe"

  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}" \
    "DisplayName" "${ProgName}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}" \
    "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegDword HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}" \
    "NoModify" 1
  WriteRegDword HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}" \
    "NoRepair" 1
SectionEnd

Section "External resources"
  SectionIn RO
  SetOutPath "$INSTDIR\externals\"
  File "${ExtFilDir}\*.*"
SectionEnd

Section "Manual"
  SetOutPath "$INSTDIR\manual\"
  File "${ManFilDir}*.*"
  SetOutPath "$INSTDIR\manual\gfx"
  File "${ManFilDir}\gfx\*.*"
SectionEnd

Section "Startmenu icons"
  SetShellVarContext all

  CreateDirectory "$SMPROGRAMS\DaisyWare"
  CreateDirectory "$SMPROGRAMS\DaisyWare\${ProgName}\"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\${ProgName}.lnk" \
    "$INSTDIR\d202validator.exe"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\${ProgNameUiIterator}.lnk" \
    "$INSTDIR\IteratorValidator.exe"

  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\License.lnk" \
    "$INSTDIR\lgpl.txt"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Readme.lnk" \
    "$INSTDIR\readme.txt"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\User Manual.lnk" \
    "$INSTDIR\manual\validator_user_manual.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Developer Manual.lnk" \
    "$INSTDIR\manual\validator_developer_manual.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\UnInstall.lnk" \
    "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Desktop icon"
  CreateShortcut "$DESKTOP\${ProgName}.lnk" "$INSTDIR\d202validator.exe"
  CreateShortcut "$DESKTOP\${ProgName}.lnk" "$INSTDIR\IteratorValidator.exe"
SectionEnd

UninstallText "Do you want to remove ${ProgName}?"

ShowUninstDetails show

UninstPage uninstconfirm
UninstPage instfiles

Section Uninstall
  RmDir /r "$INSTDIR"
  RmDir "$PROGRAMFILES\DaisyWare"
  
  SetShellVarContext all
  
  Delete "$SMPROGRAMS\DaisyWare\${ProgName}\*.*"
  RmDir "$SMPROGRAMS\DaisyWare\${ProgName}"
  RmDir "$SMPROGRAMS\DaisyWare"
  
  Delete "$DESKTOP\${ProgName}.lnk"
  
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}"

  Push "$SYSDIR\DaisyWare\dtdparser.dll"
  Call un.RemoveSharedDll

  Push "$SYSDIR\DaisyWare\validatorengine.dll"
  Call un.RemoveSharedDll

  call un.UninstallVb6Runtimes
SectionEnd

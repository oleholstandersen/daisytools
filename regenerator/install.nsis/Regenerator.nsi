!define VBFILESDIR "..\bin\vbruntimes"

!include "extras.nsh"
!include "vbsections.nsh"
!include "xmlcheck.nsh"

!Define ProgName "Daisy 2.02 Regenerator"
!Define ProgVer "20050419"

OutFile "Daisy.2.02.Regenerator.${ProgVer}.Install.exe"
Name "${ProgName} ${ProgVer} Installer"

LicenseData "..\bin\ui\gpl.txt"
LicenseText "The ${ProgName} is under the following license, please press \
  'I agree' to agree to the licensetext." "I agree"

ShowInstDetails show

!Define BinDirUi "..\bin\ui\"
!Define BinDirDll "..\bin\dll\"
!Define AudTmpDir "..\bin\ui\audiotemplates\"
!Define ResFilDir "..\bin\ui\resources\"
!Define SysFilDir "..\bin\sysfiles\"
!Define ManFilDir "..\bin\ui\manual\"
!Define TdyDocDir "..\bin\ui\tidylib\"
!Define TidyDir "..\bin\tidy\"

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
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{69A119D5-A681-4ECB-A161-F6FC35E5DBC3}" \
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

  DeleteRegValue HKLM "SOFTWARE\DaisyWare\Regenerator\Log" "Save log"
  DeleteRegValue HKCU "SOFTWARE\DaisyWare\Regenerator\Log" "Save log"

  SetOutPath "$SYSDIR"

  !define UPGRADEDLL_NOREGISTER
  !insertmacro UpgradeDll "${SysFilDir}\comdlg32.ocx" "$SYSDIR\comdlg32.ocx" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${SysFilDir}\richtx32.ocx" "$SYSDIR\richtx32.ocx" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${SysFilDir}\mscomctl.ocx" "$SYSDIR\mscomctl.ocx" "$SYSDIR\temp"
  !undef UPGRADEDLL_NOREGISTER

  call InstallVb6Runtimes

  SetOutPath "$INSTDIR" ; set target path for file commands below
  File "${BinDirUi}*.*" ; compile this to installer exe, and extract when installing

  CreateDirectory "$SYSDIR\DaisyWare"
  !insertmacro UpgradeDll "${BinDirDll}regeneratorengine.dll" \
    "$SYSDIR\DaisyWare\regeneratorengine.dll" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${BinDirDll}dtdparser.dll" \
     "$SYSDIR\DaisyWare\dtdparser.dll" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${TidyDir}tidyatl.dll" \
     "$SYSDIR\DaisyWare\tidyatl.dll" "$SYSDIR\temp"

  regDll "$SYSDIR\DaisyWare\regeneratorengine.dll"
  regDll "$SYSDIR\DaisyWare\dtdparser.dll"
  regDll "$SYSDIR\DaisyWare\tidyatl.dll"

  Push "$SYSDIR\DaisyWare\regeneratorengine.dll"
  Call AddSharedDll
  Push "$SYSDIR\DaisyWare\dtdparser.dll"
  Call AddSharedDll
  Push "$SYSDIR\DaisyWare\tidyatl.dll"
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

  SetOutPath "$INSTDIR\AudioTemplates\"
  File "${AudTmpDir}\*.*"
  
  SetOutPath "$INSTDIR\Resources\"
  File "${ResFilDir}\*.*"
  
  SetOutPath "$INSTDIR\Tidylib\"
  File "${TdyDocDir}\*.*"
SectionEnd

Section "Manual"
  SetOutPath "$INSTDIR\Manual\"
  File "${ManFilDir}*.*"
  SetOutPath "$INSTDIR\Manual\Gfx"
  File "${ManFilDir}\Gfx\*.*"
SectionEnd

Section "Startmenu icons"
  SetShellVarContext all

  CreateDirectory "$SMPROGRAMS\DaisyWare"
  CreateDirectory "$SMPROGRAMS\DaisyWare\${ProgName}\"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\${ProgName}.lnk" \
    "$INSTDIR\regenerator.exe"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Audio name reverter.lnk" \
    "$INSTDIR\regenerator_anr.exe"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\License.lnk" \
    "$INSTDIR\gpl.txt"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Readme.lnk" \
    "$INSTDIR\readme.txt"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\User Manual.lnk" \
    "$INSTDIR\manual\regenerator_manual.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Developer Manual.lnk" \
    "$INSTDIR\manual\regenerator_developer.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\UnInstall.lnk" \
    "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Desktop icon"
  CreateShortcut "$DESKTOP\${ProgName}.lnk" "$INSTDIR\regenerator.exe"
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

  Push "$SYSDIR\DaisyWare\regeneratorengine.dll"
  Call un.RemoveSharedDll

  Push "$SYSDIR\DaisyWare\tidyatl.dll"
  Call un.RemoveSharedDll

  call un.UninstallVb6Runtimes
SectionEnd

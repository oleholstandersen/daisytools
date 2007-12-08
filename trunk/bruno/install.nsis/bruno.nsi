!define VBFILESDIR "..\bin\vbruntimes"

!include "extras.nsh"
!include "vbsections.nsh"
!include "xmlcheck.nsh"

!Define ProgName "Bruno"
!Define ProgVer "20070910"

OutFile "${ProgName}.${ProgVer}.Install.exe"
Name "${ProgName} ${ProgVer} Installer"

LicenseData "..\license.txt"
LicenseText "${ProgName} is under the following license, please press \
  'i agree' to agree to the licensetext." "I agree"

ShowInstDetails show

!Define SysFilDir "..\bin\sysfiles\"
!Define BinDir "..\bin\core\"

ComponentText "Please select which components you want installed."
InstallDir "$PROGRAMFILES\DaisyWare\${ProgName}\"
DirText "Select the directory to install ${ProgName} into"

Function .onInit
LookForNSIS:
  ReadRegStr $R1 HKLM \
    "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\${ProgName}\" \
    "UninstallString"

  strCmp $R1 "" PrevNSISNotFound

  MessageBox MB_YESNO "A version of ${ProgName} is already installed. It is STRONGLY recommended to uninstall the previous version first. Quit install?" IDNO PrevNSISNotFound
  Quit

PrevNSISNotFound:
  call IsMSXML40Installed
FunctionEnd

Function .onInstSuccess
/*
  MessageBox MB_YESNO "This program requires MDAC 2.8 to be installed, press OK to launch installer." IDNO SkipMDAC
  ExecWait "$TEMP\MDAC_TYP.EXE"
SkipMDAC:
  Delete "$TEMP\MDAC_TYP.EXE"
*/
FunctionEnd

Page license
Page components
page directory
page instfiles

Section "Core files"
  SectionIn RO  ; RO means Read Only

  setOverwrite ifnewer

  SetOutPath "$SYSDIR"

  !define UPGRADEDLL_NOREGISTER
  !insertmacro UpgradeDll "${SysFilDir}\comdlg32.ocx" "$SYSDIR\comdlg32.ocx" "$SYSDIR\temp"
  !insertmacro UpgradeDll "${SysFilDir}\mscomctl.ocx" "$SYSDIR\mscomctl.ocx" "$SYSDIR\temp"

  call InstallVb6Runtimes

  SetOutPath "$INSTDIR" ; set target path for file commands below
  File /r "${BinDir}*.*" ; compile this to installer exe, and extract when installing
/*
  SetOutPath "$TEMP"
  File "..\MDAC_TYP.EXE"
*/
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

Section "Startmenu icons"
  SetShellVarContext all

  CreateDirectory "$SMPROGRAMS\DaisyWare"
  CreateDirectory "$SMPROGRAMS\DaisyWare\${ProgName}\"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\${ProgName}.lnk" \
    "$INSTDIR\bruno.exe"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\User Manual.lnk" \
    "$INSTDIR\externals\manual\bruno_user_manual.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Driver Documentation.lnk" \
    "$INSTDIR\externals\manual\bruno_driver.html"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\Readme.lnk" \
    "$INSTDIR\readme.txt"
  CreateShortcut "$SMPROGRAMS\DaisyWare\${ProgName}\UnInstall.lnk" \
    "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Desktop icon"
  CreateShortcut "$DESKTOP\${ProgName}.lnk" "$INSTDIR\bruno.exe"
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

  call un.UninstallVb6Runtimes
SectionEnd

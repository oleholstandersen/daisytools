; TPB Reader RC3 setup V0.1 - Visual Basic 6 runtime installer header script
;
; This file contains two functions whos purposes are to install/uninstall runtime files
; needed for Visual Basic 6 programs

;!define VBFILESDIR "..\VB6Binaries"
;# or
;#!define VBFILESDIR C:\Path\to\where\vbrun60sp5.exe\extracted
;Section "VB6 Runtime" VBINSTSEC
;  SectionIn RO
Function InstallVb6Runtimes

  !insertmacro UpgradeDLL ${VBFILESDIR}\Asycfilt.dll $SYSDIR\Asycfilt.dll "$TEMP"
  !insertmacro UpgradeDLL ${VBFILESDIR}\Comcat.dll $SYSDIR\Comcat.dll "$TEMP"
  !insertmacro UpgradeDLL ${VBFILESDIR}\Msvbvm60.dll $SYSDIR\Msvbvm60.dll "$TEMP"
  !insertmacro UpgradeDLL ${VBFILESDIR}\Oleaut32.dll $SYSDIR\Oleaut32.dll "$TEMP"
  !insertmacro UpgradeDLL ${VBFILESDIR}\Olepro32.dll $SYSDIR\Olepro32.dll "$TEMP"
  !define UPGRADEDLL_NOREGISTER
    !insertmacro UpgradeDLL ${VBFILESDIR}\Stdole2.tlb $SYSDIR\Stdole2.tlb "$TEMP"
  !undef UPGRADEDLL_NOREGISTER
;  # skip shared count increasing if already done once for this application
  IfFileExists $INSTDIR\myprog.exe skipAddShared
    Push $SYSDIR\Asycfilt.dll
    Call AddSharedDLL
    Push $SYSDIR\Comcat.dll
    Call AddSharedDLL
    Push $SYSDIR\Msvbvm60.dll
    Call AddSharedDLL
    Push $SYSDIR\Oleaut32.dll
    Call AddSharedDLL
    Push $SYSDIR\Olepro32.dll
    Call AddSharedDLL
    Push $SYSDIR\Stdole2.tlb
    Call AddSharedDLL
  skipAddShared:
;SectionEnd
FunctionEnd

;Section Uninstall
Function un.UninstallVb6Runtimes
  Push $SYSDIR\Asycfilt.dll
  Call un.RemoveSharedDLL
  Push $SYSDIR\Comcat.dll
  Call un.RemoveSharedDLL
  Push $SYSDIR\Msvbvm60.dll
  Call un.RemoveSharedDLL
  Push $SYSDIR\Oleaut32.dll
  Call un.RemoveSharedDLL
  Push $SYSDIR\Olepro32.dll
  Call un.RemoveSharedDLL
  Push $SYSDIR\Stdole2.tlb
  Call un.RemoveSharedDLL
;SectionEnd
FunctionEnd


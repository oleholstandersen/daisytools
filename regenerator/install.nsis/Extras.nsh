 ; AddSharedDLL
 ;
 ; Increments a shared DLLs reference count.
 ; Use by passing one item on the stack (the full path of the DLL).
 ;
 ; Usage:
 ;   Push $SYSDIR\myDll.dll
 ;   Call AddSharedDLL
 ;
 Function AddSharedDLL
   Exch $R1
   Push $R0
   ReadRegDword $R0 HKLM Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
   IntOp $R0 $R0 + 1
   WriteRegDWORD HKLM Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1 $R0
   Pop $R0
   Pop $R1
 FunctionEnd
 
 ; un.RemoveSharedDLL
 ;
 ; Decrements a shared DLLs reference count, and removes if necessary.
 ; Use by passing one item on the stack (the full path of the DLL).
 ; Note: for use in the main installer (not the uninstaller), rename the
 ; function to RemoveSharedDLL.
 ;
 ; Usage:
 ;   Push $SYSDIR\myDll.dll
 ;   Call un.RemoveSharedDLL
 ;
 Function un.RemoveSharedDLL
   Exch $R1
   Push $R0
   ReadRegDword $R0 HKLM Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
   StrCmp $R0 "" remove
     IntOp $R0 $R0 - 1
     IntCmp $R0 0 rk rk uk
     rk:
       DeleteRegValue HKLM Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1
     goto Remove
     uk:
       WriteRegDWORD HKLM Software\Microsoft\Windows\CurrentVersion\SharedDLLs $R1 $R0
     Goto noremove
   remove:
     UnRegDll $R1
     Delete /REBOOTOK $R1
   noremove:
   Pop $R0
   Pop $R1
 FunctionEnd
 
  ; Macro - Upgrade DLL File
 ; Written by Joost Verburg
 ; ------------------------
 ;
 ; Parameters:
 ; LOCALFILE   - Location of the new DLL file (on the compiler system)
 ; DESTFILE    - Location of the DLL file that should be upgraded (on the user's system)
 ; TEMPBASEDIR - Directory on the user's system to store a temporary file when the system has
 ;               to be rebooted.
 ;               For Win9x support, this should be on the same volume as the DESTFILE!
 ;               The Windows temp directory could be located on any volume, so you cannot use
 ;               this directory.
 ;
 ; Note: If you want to support Win9x, you can only use short filenames (8.3).
 ;
 ; Example of usage:
 ; !insertmacro UpgradeDLL "dllname.dll" "$SYSDIR\dllname.dll" "$SYSDIR"
 ;
 ; !define UPGRADEDLL_NOREGISTER if you want to upgrade a DLL that cannot be registered

 !macro UpgradeDLL LOCALFILE DESTFILE TEMPBASEDIR

   Push $R0
   Push $R1
   Push $R2
   Push $R3

   ;------------------------
   ;Check file and version

   IfFileExists "${DESTFILE}" "" "copy_${LOCALFILE}"

   ClearErrors
     GetDLLVersionLocal "${LOCALFILE}" $R0 $R1
     GetDLLVersion "${DESTFILE}" $R2 $R3
   IfErrors "upgrade_${LOCALFILE}"

   IntCmpU $R0 $R2 "" "done_${LOCALFILE}" "upgrade_${LOCALFILE}"
   IntCmpU $R1 $R3 "done_${LOCALFILE}" "done_${LOCALFILE}" "upgrade_${LOCALFILE}"

   ;------------------------
   ;Let's upgrade the DLL!

   SetOverwrite try

   "upgrade_${LOCALFILE}:"
     !ifndef UPGRADEDLL_NOREGISTER
       ;Unregister the DLL
       UnRegDLL "${DESTFILE}"
     !endif

   ;------------------------
   ;Try to copy the DLL directly

   ClearErrors
     StrCpy $R0 "${DESTFILE}"
     Call ":file_${LOCALFILE}"
   IfErrors "" "noreboot_${LOCALFILE}"

   ;------------------------
   ;DLL is in use. Copy it to a temp file and Rename it on reboot.

   GetTempFileName $R0
     Call ":file_${LOCALFILE}"
   Rename /REBOOTOK $R0 "${DESTFILE}"

   ;------------------------
   ;Register the DLL on reboot

   !ifndef UPGRADEDLL_NOREGISTER
     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\RunOnce" \
     "Register ${DESTFILE}" '"$SYSDIR\rundll32.exe" "${DESTFILE}",DllRegisterServer'
   !endif

   Goto "done_${LOCALFILE}"

   ;------------------------
   ;DLL does not exist - just extract

   "copy_${LOCALFILE}:"
     StrCpy $R0 "${DESTFILE}"
     Call ":file_${LOCALFILE}"

   ;------------------------
   ;Register the DLL

   "noreboot_${LOCALFILE}:"
     !ifndef UPGRADEDLL_NOREGISTER
       RegDLL "${DESTFILE}"
     !endif

   ;------------------------
   ;Done

   "done_${LOCALFILE}:"

   Pop $R3
   Pop $R2
   Pop $R1
   Pop $R0

   ;------------------------
   ;End

   Goto "end_${LOCALFILE}"

   ;------------------------
   ;Called to extract the DLL

   "file_${LOCALFILE}:"
     File /oname=$R0 "${LOCALFILE}"
     Return

   "end_${LOCALFILE}:"

  ;------------------------
  ;Set overwrite flag back

;  SetOverwrite lastused

 !macroend

 ; This function takes a whole path+filename and removes the filename
; Usage:
; push [path+filename]
; call removeFilename
; pop [pathonly]
;
; Destroys: $9
function removeFilename
  pop $9
  push $R0
  push $R1
  push $R2
  strCpy $R0 $9       ; Put $9 in $R0

  strLen $R4 $R0

  push $R4
  pop $R1
  intOp $R1 $R1 - 1
Again:
  strCpy $R2 $R0 1 $R1

  strCmp $R2 "\" slash
  intOp $R1 $R1 - 1
  intCmp 0 $R1 fail
  goto Again
Slash:
  IntOp $R1 $R1 + 1
  IntOp $R3 $R3 - $R1
  strCpy $9 $R0 $R1

  pop $R2
  pop $R1
  pop $R0
  push $9
  return

fail:
  pop $R2
  pop $R1
  pop $R0
  push ""
functionEnd

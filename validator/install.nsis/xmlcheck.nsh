 ; IsMSXML40Installed
 ;
 ; Returns on top of stack
 ; 0 (MSXML40 is not installed)
 ; or
 ; 1 (MSXML40 is installed)
 ;
 ; Usage:
 ;   Call IsMSXML40Installed
 ;   Pop $R0
 ;   ; $R0 at this point is "1" or "0"

Function IsMSXML40Installed
  Push $R0
  ClearErrors
  ReadRegStr $R0 HKCR "CLSID\{88D969C0-F192-11D4-A65F-0040963251E5}" ""

;/*  
;  MessageBox MB_OK $R0
;  IfErrors lbl_na
;    StrCpy $R0 1
;  Goto lbl_end
;  lbl_na:
;    StrCpy $R0 0
;  lbl_end:
;  Exch $R0
;
;  IntCmp $R0 1 done
;*/

  StrCmp $R0 "XML DOM Document 4.0" done
  MessageBox MB_OK "Please note that you need to have Microsoft MSXML 4 sp1 or later installed to run this application."
done:
FunctionEnd

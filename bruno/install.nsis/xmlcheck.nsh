
Function IsMSXML40Installed
  Push $R0
  ClearErrors
  ReadRegStr $R0 HKCR "CLSID\{88D969C0-F192-11D4-A65F-0040963251E5}" ""

  StrCmp $R0 "XML DOM Document 4.0" done
  MessageBox MB_OK "Please note that you need to have Microsoft MSXML 4 sp1 or later installed to run this application."
done:
FunctionEnd

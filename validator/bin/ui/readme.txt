Readme file for Daisy 2.02 Validator v1.0

GENERAL INFO
-----------------------------------------------------------
This is a validator for Daisy 2.02 digital talking books (DTB).

RELEASE COMMENTS
-----------------------------------------------------------
VERSION 20091009 (Engine version 1.0.15, Main Interface version 1.0.18)
* Fixed a bug occurring when validating DTBs longer than 99 hours using full mode.

-----------------------------------------------------------
VERSION 20080402 (Engine version 1.0.15, Main Interface version 1.0.18)
* Fixed a bug occurring when using full mode on DTBs longer than 99 hours, reported by David Gordon, RNIB.
* Fixed erroneous desktop shortcut (the Validator shortcut was pointing to ValidatorIterator executable)

-----------------------------------------------------------
VERSION 20050401 (Engine version 1.0.15, Main Interface version 1.0.18)
* Files referenced via url() statements in CSS are now detected
* Reinserted test for first h1 of NCC having @class='title' - 
  this test has mistakenly been omitted since a prior release.
* Audiofile integrity (isValidAudioFile, hasValidExtension, hasValidName, hasRecommendedName) tests 
  now also run in lightmode.
* See also note on VERSION 20041222 below. 

-----------------------------------------------------------
VERSION 20041222 (Engine version 1.0.10, Main Interface version 1.0.13)
* Added fix for non-admin user account registry handling error in 
  Regenerator reported by Heiko Becker, DZB. 
  All registry settings are now stored per user account (HKEY_CURRENT_USER).
  Note - this means that after install of this new version, local configuration settings 
  will have to be reset.
  
-----------------------------------------------------------
VERSION 20040804 (Engine version 1.0.10, Main Interface version 1.0.12)
* Second user interface: Validator Iterator, used for batch processing.
* Better checking that MSXML4 is installed. 
  If it is not, the program will not run at all, 
  avoiding bogus reports.
* Fixed bug erroneous invalidity report of content document.
  Occured when the DTD internal subset extension
  (used for skippable DTBs) was present.
  
 
VERSION 20031118 (Engine version 1.0.8, Interface version 1.0.8)
* fixed 0 byte audio files being created when an audiofile was missing
* clip times and audiofile validity no longer tested when audiofile is missing
* improved the suggested value report for smil clock values
* user interface: when saving a document, backup files now go to sub directory "val_bkp"
* user interface: document editor goes read only when doc opened is unicode
* extensions to the user manual

RC2
* removed critical/non critical branches from UI
* fix for relative paths in SMIL files
* malformed SMIL files are no longer reported as missing
* inserted PID check
* renewed code for MPEG detection
* inserted test for duplicated TEXT events within PAR elements
* fixed error with whitespace in DOCTYPE
* added 'light mode' which makes a validation pass that only
  incorporates vital tests such as file existance, wellformedness,
  DTD validity, link integrity.
* added 'disable audio tests mode' which does all tests in full
  conformance testing mode, except those concerning audio files.
* fixed a bug where ncc:depth '7' was expected when found
  and correct value was '6'
* fixed a bug in the search replace dialog
* fixed an internal error when content doc linkback href base
  (the smilfile) did not exist
* fixed an error in aux file testing where "missing file "
  was reported on file residing in relative subdir
  ("src='images/me.jpg'")
* fixed ncc:revision and ncc:revisionDate who by mistake
  was typed as optional-recommended in ncc.adtd. Now warnings
  will no more occur when these are missing in NCC.

[see also known issues below]

RC1
* VTM has been updated
* Advanced ADTD info is now optional
* A whole lot of updates in the ADTD:s, DTD:s and code
* Audiofile problems have been fixed.

Beta 2
* A whole rewrite of the program has been done since BETA 1


KNOWN ISSUES
-----------------------------------------------------------
* The procentual progress of validation is not always
  accurate.
* The line/column reported by the validator is not always correct.
* Problems may occur if validating DTBs using UTF-16.
* The validator crashes while validating certain mp3 files.
  This problem seems to occur due to a bug in Windows
  media player (DirectShow).
* Unicode encoded files will show up as having garbage
  characters when opened in the internal
  edit window. Users who validate unicode encoded books are
  adviced to not edit the files within the Validator UI.
* mp3 files encoded using VBR will not issue a warning

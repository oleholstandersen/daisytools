Readme file for Daisy 2.02 Regenerator

GENERAL INFO
-----------------------------------------------------------
This is an upgrade and fix tool for Daisy 2.x digital talking books (DTB).

RELEASE COMMENTS version 20050419 (Engine version 1.0.28, Interface version 1.0.12)
- Added a fix for a bug related to audio file references: 
  when sequential rename was run multiple times using identical prefix each time, 
  and mp3 files where inserted inbetween runs (i.e without cleanup), this would 
  result in SMIL referencing audio erroneously. (reported by Brandon, CNIB)
- Added feature per request: when external metadata import is active, ncc:narrator is kept 
  as in original ncc, unless it is present in metadata import document.

RELEASE COMMENTS version 20050403 (Engine version 1.0.27, Interface version 1.0.11)
-----------------------------------------------------------
- Now correctly identifies and brings along files referenced via url() statements in CSS.
- Removed possibility of ID case inconsistency (could occur when smil destination fragment was at <par> instead of <text>, and modify smil target set to "no change")
- GUI: fixed error dialog "property settings cancelled" when editing paths after having used "set all jobs to these settings"
- Added case correction on class attribute value on first h1 in NCC.
- Fixed bug in related to using path variables when moving books (Guillaume DuBourget)
- Added fix in registry handling: 
  unable to recall use of validator settings when not lagged in as Admin (Heiko Becker, DZB).
  
  All registry settings are now stored per user account (HKEY_CURRENT_USER).  
  Note: when installing this version of the Regenerator, also reinstall the latest version of the Validator.
  Before running the Regenerator, open and close the dedicated Validator UI. 
  This will set the new registry entries for the Validator that this version of the Regenerator needs.

- Implementors note: DLL Constant typo DTB_AUDIOFULLTTEXT corrected to DTB_AUDIOFULLTEXT.

RELEASE COMMENTS version 20041029 (Engine version 1.0.21, Interface version 1.0.6)
-----------------------------------------------------------
- Added a fix for Victor problem: when the string "id" occurs
  as the two first characters in the last word of chapter text, Victor may reset.
  Problem is caused by SMIL title metadata, and the fix done by Regenerator is
  to insert a space character at the very end of chapter title text in
  SMIL metadata.
  Visuaide comments (2004-10-28):
  "The Vibe is unaffected by the problem. For the PRO/Classic/Classic+ a
  fix is ready and a version will be available in January. A beta will be
  available for organisations interested in early December. The schedule for
  VR Soft has yet to be determined."

- Added forced remove of ncc and contentdoc metadata when malformed or erratic 
  as occured in certain early versions of SigtunaDAR3 and LpStudioPlus 
  (this caused Regenerator to abort).
  Examples of erratic state:
  <meta name="dc:publisher" content="Albert Bonniers förlag" scheme="Albert Bonniers förlag" name=dc:publishe"/>
  <meta name="dc:identifier" content="91-0-057708-1" scheme="content=91-0-057708-1 name=dc:identifier"/>

  The forced remove only happens when preserve bibliographic metadata is set to false.

- Fixed regenerator abort when rendermode was set to replace and a write protected master.smil 
  existed in original DTB directory.


RELEASE COMMENTS version 20040804 (Engine version 1.0.18, Interface version 1.0.5)
-----------------------------------------------------------
-Fixed so that DTD internal subset extensions of content documents are maintained (applies to skippability)
-Change in naming of backup and unref directories; now include the tilde char to allow exclusion by popular ISO-maker software etc.
-Various other minor fixes.

RELEASE COMMENTS version 20040330 (Engine version 1.0.13, Interface version 1.0.2)
-----------------------------------------------------------
-Broken link estimation routine enhanced.
-Better handling of DTBs made with the Sigtuna DAR 2 software.


RELEASE COMMENTS version 20031118 (Engine version 1.0.3, Interface version 1.0.2)
-----------------------------------------------------------
-Bookfix routine enhanced.
-Several bugfixes.
-Added new optional functionality:
 - redirect SMIL targets to par/text;
 - make true ncc only;
 - estimation and reinsertion of broken links;
 - add default css;
 - log verbosity.
-Read more about these functions in the manual.


RELEASE COMMENTS Beta 1
-----------------------------------------------------------
First public release.


KNOWN ISSUES
-----------------------------------------------------------
In the batch user interface, validation cannot be enabled unless
the validator has been previously run through its own user interface.
Readme file for Bruno, Daisy Fileset creator

RELEASE NOTES VERSION 20070910 (exe version 0.9.79)
------------------------------
- remade the 'fixed omitted sync' fix from previous release as it caused other unwanted behavior in mixed content models with interspersed text, sync-include and sync-omit elements.
- fixed correct sync on parent when the parent only contains sync-omit children.

RELEASE NOTES VERSION 20070531 (exe version 0.9.78)
------------------------------
- fixed omitted sync when more than one emphasis or strong occured with other synced siblings (driver mod only)
- fixed a "stop statement encountered" poof reported by Brandon Nelson, CNIB.
- various fixes to the Z39.86-2005 rendering as requested by various people
- added feature: default XHTML driver supporting representation in NCC of full set of skippable elements
- added fix/feature: occurences of div.prodnode and div.sidebar in content document turned into span in ncc
- added fix: mixed content nodes no longer loose text in ncc (ie "<h1> my <em>inline</em> text</h1>" before fix became "<h1> my  text</h1>" in ncc)
- added fix: bodyref attrs removed from ncc

Note - Dtbook-to-Z3986.2005 mode is vastly improved but still experimental; output fileset is not guaranteed to be bug free. Only use the XHTML-to-Daisy 2.02 mode if you are in sharp a production environment.

RELEASE NOTES VERSION 20051214 (exe version 0.9.72)
------------------------------
- upgraded the ANSI/NISO support to version 2005. Note - this mode is still experimental; output fileset is not guaranteed to be bug free.


RELEASE NOTES VERSION 20050425 (exe version 0.9.71)
------------------------------
- added fix: ncc:generator in SMIL being null after opening project in lpPro (David Gordon RNIB)
- added fix: pretty print routine that removed alphanumeric entities in XHTML (Brandon Nelson CNIB)
- added fix: import into lpp and ncc of additional meta items in XHTML sourcedoc (Niels Thogersen IBOS)
- added fix: SQL error when inserting metadata containing single quotes (Sean Brooks CNIB)
- added better handling when auxilliary files (css, images) resides in subdirectories of input document
- added a new mdf that enables heading level editing within LpStudioPro (provided by Sean Brooks CNIB)

RELEASE NOTES VERSION 20041222 (exe version 0.9.65)
------------------------------
-  added fix: runtime error when no namespace declaration present (Miki Azuma)
-  added feature: support of url() statements in CSS for inclusion of auxilliary files (DFA ITT)
-  added feature: working status indication in caption (Niels Thogersen)
-  added fix: handling of spaces in input and output directory paths (Sean Brooks)
-  added feature/fix: validation of XHTML now also checks for first body element being an h1 class title (Sean Brooks)
-  added fix: keeping infobar message history
-  change in registry handling: outputpath and driver selection stored between sessions - now per user profile.

RELEASE NOTES VERSION 20041029 (exe version 0.9.62)
------------------------------
- added fix: whitespace truncation on inlines 
  (reported by Jesper Klein)
- added fix: problem in QAPlayer caused by first two headings 
  being in same SMIL file (reported by David Gordon and Stefan Kropf)
- added fix: problem in content doc display in LpPro when head/title 
  was empty or whitespace only (reported by David Gordon)
- added fix: Lpp creation error when several input documents pointed 
  to auxilliary files in different folders but with same filename 
  (reported by Per Sennels)
- added feature: outputpath and driver selection stored between 
  sessions (suggested by Per Sennels)

RELEASE NOTES VERSION 20040903
------------------------------
First public beta release
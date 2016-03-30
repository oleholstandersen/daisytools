# Daisy 2.02 Validator Known Issues #

This page lists known false negatives in the Daisy 2.02 Validator output report.

In other words, this means that the Validator claims that an error exists in the DTB, while it is not necessarily so.

Due to resource constraints, the issues listed below are unlikely to be adressed anytime soon.

## [Issue 1](https://code.google.com/p/daisytools/issues/detail?id=1). MP3 Files and ID3 tags ##
**Issue**: MP3 Files containing certain versions of ID3 tags are flagged as corrupt by the Validator.

**Reason**: The Daisy 2.02 Validator does not properly recognize all varieties of ID3 tags.

**Solution**: When you recieve a report like this, make sure that its due to the presence
of ID3 tags by opening the file in some MP3 Player software and checking that ID3 is present.

_Note: some encoders will add an empty ID3 tag (""), this still counts as ID3!_

## [Issue 2](https://code.google.com/p/daisytools/issues/detail?id=2). File Count and File Size in Multi Volume DTBs ##
**Issue**:
Under certain conditions in Multi Volume DTBs, the Validator will miscalculate the number of expected files (meta item _ncc:files_) and the expected file size (meta item _ncc:kByteSize_).

**Reason**: There is a bug in the way the validator calculates the number of files and handles duplicates across the set of DTB volumes.

**Solution**: _ncc:files_ and  _ncc:kByteSize_ are not critical for correct playback of the DTB. If you still want to make sure that the DTB is 100% valid, manually confirm that the given number of files, and their size is correct.

## [Issue 3](https://code.google.com/p/daisytools/issues/detail?id=3). XML prolog and DOCTYPE declaration on the same line ##

**Issue**: The validator unexpectedly reports the following error:
```
error (critical): : document does not validate against the DTD referenced in the document, Invalid at the top level of the document.
?xml version="1.0" encoding="utf-8"?>
```

**Reason**:  The validator does not accept when the XML prolog (`<?xml ... ?>`) and the Doctype declaration (`<!DOCTYPE ...>`) are on the same line. This is a technical limitation (since there is/was no entity resolver available in msxml4 DOM, the prolog was "hacked" with string manipulation).

**Solution**: Insert line breaks around the DOCTYPE declaration.
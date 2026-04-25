# Changes from Original Evernote2Onenote

This fork extends [Stefan Kueng's Evernote2Onenote](https://tools.stefankueng.com/Evernote2Onenote.html) with multi-file folder import and several usability improvements.

## Changes to tag handling
Previously, sections were created for each tag in a note. This has been removed and tags now get added as a pre-pended list at the beginning of notes. (e.g. Tags: tag1, tag2). This allows searching of the terms in OneNote at least.

## New Features

### Multi-file Folder Import
The original tool imports a single `.enex` file at a time. This fork adds a **folder-based import** that processes an entire export directory in one run.

- Select an **Import folder** containing `.enex` files
- Subfolders become **section groups** in OneNote (matching Evernote's "stack" concept)
- Root-level `.enex` files always become **direct sections** in the notebook root.

### Manifest-based Progress Tracking
Each import run writes a **manifest file** to a configurable log folder. The manifest records the status of every file (pending, in-progress, complete, failed).

- If a run is interrupted, restarting the tool offers to **resume** from where it left off
- Files that fail are retried up to 3 times before being marked permanently failed

### Notebook Name Field
The target OneNote notebook name is now a field on the main form (rather than being inferred automatically). 

- Browsing for an import folder auto-suggests a notebook name based on the folder name
- If the named notebook already exists in OneNote, the tool prompts before proceeding

### Editable Path Fields
The Import folder and Manifest log folder fields accept typed or pasted paths in addition to browsed selections.

## Bug Fixes
There was a comment on Reddit (https://www.reddit.com/r/OneNote/comments/1rbafj1/has_anyone_recently_used_evernote2onenote_or/) noting that the app would crash after 3 or so imports. This seems to be related to timing around processing items in OneNote, essentially not allowing enough time for OneNote to complete one action before starting another. Adding delays seems to have corrected the issue.

## UI Improvements

- Taller form with more vertical breathing room between controls
- Label text no longer clipped by overlapping textboxes
- Browse buttons fully visible

## Note on imported data

Not particular to this updated version, but I noted in my imported data that ink notes don't display properly, as OneNote doesn't support the format (not a surprise, of course). Just something to keep in mind.

# OfficeImageExtractor

Extracts images in documents created by Office 2007+.

[中文](https://github.com/Mark9804/OfficeImageExtractor/blob/master/README.md)

## Function

Extracts images in documents created by Office 2007+ (docx, pptx, xlsx). Maybe videos can be extracted too, but I haven’t tried yet.

## How it works

Office 2007 and later version of Office created documents will save media files to `[format]\media`.  The core code of this script extracts documents as zip archives, then move the images according to file format.

\* Format-path correspondence：

| File format (07-2016) | Corresponding path |
| --------------------- | ------------------ |
| Word                  | ~\word\media       |
| PowerPoint            | ~\ppt\media        |
| Excel                 | ~\xl\media         |

## Usage

###  1. From terminal

`python OfficeImageExtractor.py relative-path-to-document`

###  2. Using an executable

Drag & drop a document onto the [executable](https://github.com/Mark9804/OfficeImageExtractor/releases).

**The script would open the destination folder once complete.** The destination folder is under the same root of the original document, and they share the same filename (but without extension).

## TODO

- [ ] Add batch file support
- [ ] Mac port\*
- [ ] Add support for Office 97-2003 doucments\*\*

\* The script uses Windows batch commands to delete and open folders, thus cannot be ported to \*Unix platforms directly.

\*\*The way Office 97-2003 documents save their media file is quite different from Office 07-2016 format. I haven’t figured out how the old format does the work.
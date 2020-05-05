# Pandoc-Common
Common files used as part of the pandoc publishing process

## Introduction
The current process to author/edit a document for the [International Organization for Standardization (ISO)](https://www.iso.org) requires the use of MSWord and a special template.  While it makes authoring easier for some things, it makes comparison & tracking of changes across revisions a manual process.  It also restricts authoring to a select few individuals who can easily share the Word file - and hopefully not step on each other's toes.

The goal of the contents of this repository is to enable authoring/publishing from Markdown.

## Details
This repository contains the necessary additions to the [Pandoc](http://pandoc.org) tool chain to take a Markdown file and convert it to MSWord for delivering into the ISO publication process.  There are three pieces of the process

- [reference.docx](reference.docx) which is a modified version of the standard ISO template.  It has been adjusted to align with the Pandoc DocX Writer export reference requirements plus additional commonly used styles.
- [WordExtrasFilter.lua](WordExtrasFilter.lua) which is a custom Lua filter for Pandoc that helps with additional functionality and mapping from the Markdown to the ISO Template doc.
- [2docx.yml](2docx.yml) is the YAML file that brings them all together when combined with a command line like:

```
pandoc --defaults 2docx.yml --no-highlight -o output.docx input.md
```

## Contacts
For more information about this work, please contact [Leonard Rosenthol](mailto:lrosenth@adobe.com) who is the chair for ISO TC 171 SC2 and the author of this material.

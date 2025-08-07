=== phpxlsx 2.5 ===
https://www.phpxlsx.com/

PHPXLSX is a PHP library designed to dynamically generate spreadsheets in Excel format (SpreadsheetML).

=== What's new on phpxlsx 2.5? ===

- New methods to insert comments: addComment and addCommentAuthor.
- New method to insert defined names: addDefinedName.
- New method to add a macro: addMacroFromXlsx.
- Indexer (Advanced and Premium licenses):
    · Comments.
    · Defined names.
    · Cell positions.
    · Cell formats.
- Enabled fullCalcOnLoad and forceFullCalc options in the default base template.
- HTML to XLSX supports transforming tables.
- Improvements for working with tables:
    · Add title and description properties.
    · Apply content types to contents and totals.
    · addTable returns the position range used.
- New options in setWorkbookSettings: readOnly, fullCalcOnLoad, forceFullCalc.
- New options in addSheet: position, color.
- New setCellValue method to set a cell value while keeping existing styles.
- The getCell method returns the style index value.
- Images can be added from files, streams and base64 contents.
- Applied htmlspecialchars to content values when using template methods.
- Improved handle XLSX templates that use uppercase ContentType extensions.
- Added '/' as delimiter in preg_quote when extracting the template variables.
- The static variables $_templateSymbolStart and $_templateSymbolEnd available in the CreateXlsxFromTemplate class have been changed to class attributes.
- Free DOMDocument resources when using Indexer (Advanced and Premium licenses).
- The overwrite option has been renamed to replace in all methods.
- Sign class avoid generating partial signed files in the latest revisions of MS Office (Premium licenses).

2.0 VERSION

- New methods to insert headers and footers: addHeader and addFooter.
- New method to insert SVG contents: addSvg.
- New method to insert CSV contents: addCsv (Advanced and Premium licenses).
- New method to create and apply cell styles: createCellStyle.
- PHP 8.2 support.
- addTable supports applying filters (values and custom).
- Added new chart types in addChart: bar3DCylinder, bar3DCone, bar3DPyramid, col3DCylinder, col3DCone, col3DPyramid, area, area3D.
- Added the hyperlink option to addImage to set hyperlinks in images.
- Options to hide contents in addSheet, setSheetSettings, setRowSettings and setColumnSettings.
- setActiveSheet allows using -1 as position to choose the last sheet.
- Indexer (Advanced and Premium licenses):
    · Headers.
    · Footers.
    · Cell styles.
- New getCellPositions method to return existing cell positions in the active sheet.
- HTML to XLSX:
    · New tags: sub, sup.
    · New styles: vertical-align (super, sub).
- New view option in setSheetSettings.
- Added support for strict tags when using image template methods. Supported the drawingHF tag.
- Updated placeholder names used internally by phpxlsx to generate XML contents adding '__PHX=' prefixes.
- getTemplateVariables doesn't return duplicated placeholder names in the same target.

1.0 VERSION

- Support for all MS Excel versions from MS Excel 2007 to MS Excel 2021. Other XLSX readers such as LibreOffice and Google Docs are supported too (the support of these programs reading contents and styles in XLSX files may vary).
- Generate XLSX files from scratch and using templates.
- Content methods: addBreak, addCell, addCellRange, addChart, addFunction, addImage, addLink, addSheet, addTable, getCell.
- Layout and general methods: addBackgroundImage, addProperties, getActiveSheet, setActiveSheet, setColumnSettings, setMarkAsFinal, setRowSettings, setRtl, setSheetSettings, setWorkbookSettings.
- Template methods: getTemplateVariables, removeVariableText, replaceVariableImage, replaceVariableText, setTemplateSymbol.
- HTML to XLSX.
- RTL support.
- Transform XLS to XLSX, XLSX to PDF, XLSX to ODS, ODS to XLSX (Advanced and Premium licenses).
- Indexer: return information from a XLSX (Advanced and Premium licenses).
- Crypto: protect and encrypt XLSX files, remove protection from XLSX files (Premium licenses).
- Sign XLSX files (Premium licenses).
- Save and download XLSX files.
- Stream mode (Premium licenses).
- XLSXUtilities: searchAndReplace, split (Advanced and Premium licenses).

=== What are the minimum technical requirements? ===
To run phpxlsx you need to have a functional PHP setup, this should include:

- PHP 5.2.11 or newer.
- A webserver (such as Apache, Nginx or Lighttpd) or PHP CLI.
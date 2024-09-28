# ExcelToWordTemplateGen

Just a simple console application that replaces placeholders from a word document (template) using the values (rows) and definitions (columns - 1st row) from an excel spreadsheet.

The config settings are read from appsettings.json:

1. InputDirectory
   - Input folder where the excel sheets go in
2. OutputDirectory
   - Output folder where the word docs are generated to
3. TemplateFilePath
   - The actual template that's getting transformed
4. StaticFileNameStart
   - Fixed first part of the filename
5. DynamicFieldNames
   - Takes the value from input row as part of filename and adds to it - if empty uses definition itself
6. DynamicFileNameDelimiter
   - Separator that separates the values of the dynamic fields above in the output filename  

```json
"Generator": {
  "InputDirectory": "F:\\Test\\Input",
  "OutputDirectory": "F:\\Test\\Output",
  "TemplateFilePath": "F:\\Test\\Template.docx",
  "Output": {
    "StaticFileNameStart": "Generated",
    "DynamicFieldNames": "Firstname;Lastname",
    "DynamicFileNameDelimiter":  "_"
  }
}
```

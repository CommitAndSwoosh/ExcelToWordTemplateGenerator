# ExcelToWordTemplateGen

Just a simple console application that replaces placeholders from a word document (template) using the values (rows) and definitions (columns - 1st row) from an excel spreadsheet.

The config settings are read from appsettings.json:

"Generator": {
  "InputDirectory": "F:\\Test\\Input", //Input folder where the excel sheets go in
  "OutputDirectory": "F:\\Test\\Output", //Output folder where the word docs are generated to
  "TemplateFilePath": "F:\\Test\\Template.docx", //The actual template that's getting transformed
  "Output": {
    "StaticFileNameStart": "Generated", //Fixed first part of the filename
    "DynamicFieldNames": "Firstname;Lastname", //Takes the value from input row as part of filename and adds to it - if empty uses definition itself
    "DynamicFileNameDelimiter":  "_" //Separator that separates the values of the dynamic fields above in the output filename
  }
}

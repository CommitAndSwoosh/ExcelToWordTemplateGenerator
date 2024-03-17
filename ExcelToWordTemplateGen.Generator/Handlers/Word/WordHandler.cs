using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToWordTemplateGen.Generator.Handlers.Word;

public class WordHandler : IWordHandler
{
    private readonly ILogger<WordHandler> _logger;
    private readonly OutputOptions _outputOptions;

    public WordHandler(ILogger<WordHandler> logger, IOptions<OutputOptions> outputOptions)
    {
        _logger = logger;
        _outputOptions = outputOptions.Value;
    }

    public List<string> GenerateWordFiles(string templateFilePath, string outputDirectory, DataTable data)
    {
        List<string> generatedFiles = [];

        foreach (DataRow dataRow in data.Rows)
        {
            string suffix = GetSuffixFileNamePart(_outputOptions.DynamicFieldNames, _outputOptions.DynamicFileNameDelimiter, data, dataRow);
            string pathName = Path.Combine(outputDirectory, _outputOptions.StaticFileNameStart + suffix);
            int existingFiles = Directory.GetFiles(outputDirectory, _outputOptions.StaticFileNameStart + suffix + "*.docx").Count();

            if (existingFiles > 0)
            {
                pathName = pathName + "_" + (existingFiles + 1);
            }

            using (var file = WordprocessingDocument.Open(templateFilePath, false))
            {
                if (file != null)
                {
                    string clonedFileName = pathName + ".docx";
                    using var clonedFile = file.Clone(clonedFileName, true);
                    if (clonedFile != null)
                    {
                        ReplacePlaceholders(clonedFile, data, dataRow);
                        clonedFile.Save();
                        generatedFiles.Add(clonedFileName);
                    }
                }
                else
                {
                    _logger.LogError($"Unable to clone file from {templateFilePath} to {pathName}");
                }
            }
        }

        return generatedFiles;
    }

    private string GetSuffixFileNamePart(string suffixDefinition, string filedNamesDelimiter, DataTable dataTable, DataRow dataRow)
    {
        if (!string.IsNullOrEmpty(suffixDefinition) && dataTable is not null)
        {
            var splits = suffixDefinition.Split(';');
            var results = new List<string>();

            string result = "";

            foreach (string split in splits)
            {
                var column = dataTable.Columns[split.ToUpper()];
                if (column != null)
                {
                    var value = dataRow[column].ToString();

                    if (string.IsNullOrEmpty(value))
                    {
                        _logger.LogInformation($"No value found for suffixDefinition, using the definition itself {split}");
                    }
                    else
                    {
                        results.Add(value);
                    }
                }
            }

            if (results.Count > 0)
                result = string.Join(filedNamesDelimiter, results);
           
            return result;
        }
        return suffixDefinition;
    }

    private void ReplacePlaceholders(WordprocessingDocument doc, DataTable dataTable, DataRow dataRow)
    {
        string? docText = null;

        if (doc.MainDocumentPart is null)
        {
            throw new ArgumentNullException("MainDocumentPart and/or Body is null.");
        }

        using (StreamReader sr = new(doc.MainDocumentPart.GetStream()))
        {
            docText = sr.ReadToEnd();
        }

        foreach (DataColumn column in dataTable.Columns)
        {
            var rowValue = dataRow[column].ToString();
            Regex regex = new("%" + column.ColumnName.ToUpper() + "%");
            docText = regex.Replace(docText, rowValue!);
        }

        using StreamWriter sw = new(doc.MainDocumentPart.GetStream(FileMode.Create));
        sw.Write(docText);
    }
}

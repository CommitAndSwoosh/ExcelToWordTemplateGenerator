using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Logging;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToWordTemplateGen.Generator.Handlers;

public class WordHandler : IWordHandler
{
    private readonly ILogger<WordHandler> _logger;
    public WordHandler(ILogger<WordHandler> logger)
    {
        _logger = logger;
    }

    public List<string> GenerateWordFiles(string templateFilePath, string outputDirectory, DataTable data, string prefix, string suffixDefinition)
    {
        List<string> generatedFiles = new List<string>();

        foreach (DataRow dataRow in data.Rows)
        {
            string suffix = GetSuffixFileNamePart(suffixDefinition, data, dataRow);
            string pathName = Path.Combine(outputDirectory, prefix + suffix);
            int existingFiles = Directory.GetFiles(outputDirectory, prefix + suffix + "*.docx").Count();

            if (existingFiles > 0)
            {
                pathName = pathName + "_" + (existingFiles + 1);
            }

            using (var file = WordprocessingDocument.Open(templateFilePath, false))
            {
                if (file != null)
                {
                    string clonedFileName = pathName + ".docx";
                    using (var clonedFile = file.Clone(clonedFileName, true))
                    {
                        if (clonedFile != null)
                        {
                            ReplacePlaceholders(clonedFile, data, dataRow);
                            clonedFile.Save();
                            generatedFiles.Add(clonedFileName);
                        }
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

    private string GetSuffixFileNamePart(string suffixDefinition, DataTable dataTable, DataRow dataRow)
    {
        if (!string.IsNullOrEmpty(suffixDefinition) && dataTable is not null)
        {
            var column = dataTable.Columns[suffixDefinition.ToUpper()];
            if (column != null)
            {
                var value = dataRow[column].ToString();

                if (string.IsNullOrEmpty(value))
                {
                    _logger.LogInformation($"No value found for suffixDefinition, using the definition itself {suffixDefinition}");
                }
                else
                {
                    return value;
                }
            }
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

        using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream()))
        {
            docText = sr.ReadToEnd();
        }

        foreach (DataColumn column in dataTable.Columns)
        {
            var rowValue = dataRow[column].ToString();
            Regex regex = new Regex("%" + column.ColumnName.ToUpper() + "%");
            docText = regex.Replace(docText, rowValue!);
        }

        using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
        {
            sw.Write(docText);
        }
    }
}

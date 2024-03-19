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

            using var file = WordprocessingDocument.Open(templateFilePath, false);
            if (file != null)
            {
                string clonedFileName = pathName + ".docx";
                using (var clonedFile = file.Clone(clonedFileName, true))
                {
                    if (clonedFile != null)
                    {
                        if (clonedFile.MainDocumentPart is null)
                            throw new ArgumentNullException("clonedFile.MainDocumentPart and/or Body is null.");

                        string? textToReplace;

                        using (StreamReader sr = new(clonedFile.MainDocumentPart.GetStream()))
                        {
                            textToReplace = sr.ReadToEnd();
                        }

                        textToReplace = ReplacePlaceholders(textToReplace, data, dataRow);

                        using StreamWriter sw = new(clonedFile.MainDocumentPart.GetStream(FileMode.Create));
                        sw.Write(textToReplace);
                        generatedFiles.Add(clonedFileName);
                    }
                }
            }
            else
            {
                _logger.LogError($"Unable to clone file from {templateFilePath} to {pathName}");
            }
        }

        return generatedFiles;
    }

    public string GetSuffixFileNamePart(string suffixDefinition, string filedNamesSeparator, DataTable dataTable, DataRow dataRow)
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
                result = string.Join(filedNamesSeparator, results);
           
            return result;
        }
        return suffixDefinition;
    }

    public string ReplacePlaceholders(string input, DataTable dataTable, DataRow dataRow)
    {
        if(dataTable is null || dataRow is null)
        {
            return input;
        }

        string output = input;

        foreach (DataColumn column in dataTable.Columns)
        {
            var rowValue = dataRow[column].ToString();
            Regex regex = new("%" + column.ColumnName.ToUpper() + "%");
            output = regex.Replace(output, rowValue!);
        }

        return output;
    }
}

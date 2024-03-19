using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers.Word;

public interface IWordHandler
{
    public List<string> GenerateWordFiles(string templateFilePath, string outputDirectory, DataTable data);
    public string GetSuffixFileNamePart(string suffixDefinition, string filedNamesDelimiter, DataTable dataTable, DataRow dataRow);
    public string ReplacePlaceholders(string input, DataTable dataTable, DataRow dataRow);
}

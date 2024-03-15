using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers;

public interface IWordHandler
{
    public List<string> GenerateWordFiles(string templateFilePath, string outputDirectory, DataTable data, string prefix, string suffixDefinition);
}

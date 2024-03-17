using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers.Word;

public interface IWordHandler
{
    public List<string> GenerateWordFiles(string templateFilePath, string outputDirectory, DataTable data);
}

using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers;

public interface IExcelHandler
{
    DataTable ReadExcelDefinitions(string filePath);
}

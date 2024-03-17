using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers.Excel;

public interface IExcelHandler
{
    DataTable ReadExcelDefinitions(string filePath);
}

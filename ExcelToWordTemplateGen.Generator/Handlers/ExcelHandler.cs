using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging;
using System.Data;

namespace ExcelToWordTemplateGen.Generator.Handlers;

public sealed class ExcelHandler : IExcelHandler
{
    private readonly ILogger<ExcelHandler> _logger;

    public ExcelHandler(ILogger<ExcelHandler> logger)
    {
        _logger = logger;
    }

    public DataTable ReadExcelDefinitions(string filePath)
    {
        var result = new DataTable();

        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = doc.WorkbookPart ?? doc.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            string? text;

            int current = 0;

            foreach (Row r in sheetData.Elements<Row>())
            {
                if (current > 0)
                {
                    var tmpRow = result.NewRow();
                    int currentCell = 0;
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c?.CellValue?.Text;
                        var cellval = GetCellValue(doc, c);
                        tmpRow[currentCell] = cellval;
                        currentCell++;
                    }
                    result.Rows.Add(tmpRow);

                }
                //Read Datatable column definitions
                else
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c?.CellValue?.Text;

                        if (text is not null)
                        {
                            var cellval = GetCellValue(doc, c);
                            result.Columns.Add(cellval.ToUpper());
                        }

                    }
                }
                current++;
            }
        }

        return result;
    }

    public static string GetCellValue(SpreadsheetDocument document, Cell cell)
    {
        SharedStringTablePart stringTablePart = document.WorkbookPart!.SharedStringTablePart!;

        if (cell is not null && cell.CellValue is not null)
        {
            string value = cell.CellValue.InnerXml;
            if (cell != null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        return "";
    }
}

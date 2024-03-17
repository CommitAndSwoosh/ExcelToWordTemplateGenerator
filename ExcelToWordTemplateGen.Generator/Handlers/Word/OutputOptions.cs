namespace ExcelToWordTemplateGen.Generator.Handlers.Word;

public class OutputOptions
{
    public string StaticFileNameStart { get; set; } = string.Empty;
    public string DynamicFieldNames { get; set; } = string.Empty;
    public string DynamicFileNameDelimiter { get; set; } = string.Empty;
}

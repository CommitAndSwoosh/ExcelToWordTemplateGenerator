namespace ExcelToWordTemplateGen.Generator;

public class GeneratorOptions
{
    public string InputDirectory { get; set; } = string.Empty;
    public string OutputDirectory { get; set; } = string.Empty;
    public string TemplateFilePath { get; set; } = string.Empty;
    public string OutputFileNamePrefix { get; set; } = string.Empty;
    public string OutputFileNameSuffixDefinition { get; set; } = string.Empty;
}

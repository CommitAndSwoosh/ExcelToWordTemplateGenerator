using ExcelToWordTemplateGen.Generator.Handlers.Excel;
using ExcelToWordTemplateGen.Generator.Handlers.Word;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Data;

namespace ExcelToWordTemplateGen.Generator;

public sealed class Generator : IGenerator
{
    private readonly GeneratorOptions _options;
    private readonly ILogger<Generator> _logger;
    private readonly IExcelHandler _excelHandler;
    private readonly IWordHandler _wordHandler;

    public Generator(IOptions<GeneratorOptions> options, IExcelHandler excelHandler, IWordHandler wordHandler, ILogger<Generator> logger)
    {
        _options = options.Value;
        _excelHandler = excelHandler;
        _wordHandler = wordHandler;
        _logger = logger;
    }

    public bool GenerateFiles()
    {
        if (Setup())
        {
            var excelFiles = Directory.GetFiles(_options.InputDirectory, "*.xlsx");
            List<DataTable> excelDefinitions = new();

            foreach (var excelFile in excelFiles)
            {
                var definitions = _excelHandler.ReadExcelDefinitions(excelFile);
                if (definitions != null && definitions.Rows.Count > 0)
                {
                    _logger.LogInformation($"Found {definitions.Rows.Count} rows to process in file '{excelFile}'");
                    excelDefinitions.Add(definitions);
                }
            }

            if (excelDefinitions.Count > 0)
            {
                foreach (var definition in excelDefinitions)
                {
                    if (definition != null)
                    {
                        var createdFileNames = _wordHandler.GenerateWordFiles(
                            _options.TemplateFilePath,
                            _options.OutputDirectory,
                            definition);

                        foreach (var fileName in createdFileNames)
                        {
                            _logger.LogInformation($"Created File: {fileName}");
                        }
                    }
                }
            }
        }

        return true;
    }

    private bool Setup()
    {
        _logger.LogInformation("Initializing Generator");

        if (_options is null)
        {
            _logger.LogError("No options provided - Setup failed.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_options.InputDirectory))
        {
            _logger.LogError("InputDirectory must be given - Setup failed.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_options.OutputDirectory))
        {
            _logger.LogError("OuputDirectory must be given - Setup failed.");
            return false;
        }

        if (!Directory.Exists(_options.InputDirectory))
        {
            _logger.LogError($"Input directory does not exist '{_options.InputDirectory}'");
            return false;
        }
        else
        {
            if (Directory.GetFiles(_options.InputDirectory).Length == 0)
            {
                _logger.LogInformation($"No files to process found in '{_options.InputDirectory}'");
            }
        }

        if (!Directory.Exists(_options.OutputDirectory))
        {
            _logger.LogWarning($"Output directory does not exist, trying to create '{_options.OutputDirectory}'");

            try
            {
                var directory = Directory.CreateDirectory(_options.OutputDirectory);

                if (directory.Exists)
                {
                    _logger.LogInformation($"Output directory has been successfully created '{_options.OutputDirectory}'");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Unable to create directory {_options.OutputDirectory}");
            }
        }

        return true;
    }
}

using ExcelToWordTemplateGen.Generator.Handlers;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Data;

namespace ExcelToWordTemplateGen.Generator;

public sealed class Generator : IGenerator
{
    private readonly IOptions<GeneratorOptions> _options;
    private readonly ILogger<Generator> _logger;
    private readonly IExcelHandler _excelHandler;
    private readonly IWordHandler _wordHandler;

    public Generator(IOptions<GeneratorOptions> options, IExcelHandler excelHandler, IWordHandler wordHandler, ILogger<Generator> logger)
    {
        _options = options;
        _excelHandler = excelHandler;
        _wordHandler = wordHandler;
        _logger = logger;
    }

    public bool GenerateFiles()
    {
        if (Setup())
        {
            var excelFiles = Directory.GetFiles(_options.Value.InputDirectory, "*.xlsx");
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
                            _options.Value.TemplateFilePath,
                            _options.Value.OutputDirectory,
                            definition,
                            _options.Value.OutputFileNamePrefix,
                            _options.Value.OutputFileNameSuffixDefinition);

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

        if (string.IsNullOrWhiteSpace(_options.Value.InputDirectory))
        {
            _logger.LogError("InputDirectory must be given - Setup failed.");
            return false;
        }

        if (string.IsNullOrWhiteSpace(_options.Value.OutputDirectory))
        {
            _logger.LogError("OuputDirectory must be given - Setup failed.");
            return false;
        }

        if (!Directory.Exists(_options.Value.InputDirectory))
        {
            _logger.LogError($"Input directory does not exist '{_options.Value.InputDirectory}'");
            return false;
        }
        else
        {
            if (Directory.GetFiles(_options.Value.InputDirectory).Length == 0)
            {
                _logger.LogInformation($"No files to process found in '{_options.Value.InputDirectory}'");
            }
        }

        if (!Directory.Exists(_options.Value.OutputDirectory))
        {
            _logger.LogWarning($"Output directory does not exist, trying to create '{_options.Value.OutputDirectory}'");

            try
            {
                var directory = Directory.CreateDirectory(_options.Value.OutputDirectory);

                if (directory.Exists)
                {
                    _logger.LogInformation($"Output directory has been successfully created '{_options.Value.OutputDirectory}'");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Unable to create directory {_options.Value.OutputDirectory}");
            }
        }

        return true;
    }
}

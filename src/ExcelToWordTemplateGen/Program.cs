// See https://aka.ms/new-console-template for more information

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using ExcelToWordTemplateGen.Generator;
using ExcelToWordTemplateGen.Generator.Handlers.Word;
using ExcelToWordTemplateGen.Generator.Handlers.Excel;

var host = AppStartup(args);
var scope = host.Services.CreateScope();
var generator = scope.ServiceProvider.GetRequiredService<IGenerator>();

if (generator is not null)
{
    generator.GenerateFiles();
}
else
{
    Log.Logger.Warning("Unable to get required generator service");
}

static void ConfigurationSetup(IConfigurationBuilder builder)
{
    builder.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .AddEnvironmentVariables();
}

static IHost AppStartup(string[] args)
{
    var builder = new ConfigurationBuilder();
    ConfigurationSetup(builder);

    Log.Logger = new LoggerConfiguration()
        .Enrich.FromLogContext()
        .ReadFrom.Configuration(builder.Build())
        .CreateLogger();

    var host = Host.CreateDefaultBuilder(args)
        .ConfigureServices((context, services) =>
        {
            services.Configure<GeneratorOptions>(context.Configuration.GetSection("Generator"));
            services.Configure<OutputOptions>(context.Configuration.GetSection("Generator:Output"));
            services.AddScoped<IGenerator, Generator>();
            services.AddScoped<IWordHandler, WordHandler>();
            services.AddScoped<IExcelHandler, ExcelHandler>();
        })
        .UseSerilog()
        .Build();

    return host;
}
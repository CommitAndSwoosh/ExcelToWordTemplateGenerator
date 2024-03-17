// See https://aka.ms/new-console-template for more information

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using ExcelToWordTemplateGen.Generator;
using ExcelToWordTemplateGen.Generator.Handlers;

var application = AppStartup();
//application.Run();

var generator = application.Services.GetService<IGenerator>();

if (generator is not null)
{
    generator.GenerateFiles();
}
else
{
    Log.Logger.Warning("Unable to get required generator service");
}

await application.StopAsync();


static void ConfigurationSetup(IConfigurationBuilder builder)
{
    builder.SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
        .AddEnvironmentVariables();
}

static IHost AppStartup()
{
    var builder = new ConfigurationBuilder();
    ConfigurationSetup(builder);

    Log.Logger = new LoggerConfiguration()
        .Enrich.FromLogContext()
        .ReadFrom.Configuration(builder.Build())
        .CreateLogger();

    var host = Host.CreateDefaultBuilder()
        .ConfigureServices((context, services) =>
        {
            services.Configure<GeneratorOptions>(context.Configuration.GetSection("Generator"));
            services.AddScoped<IGenerator, Generator>();
            services.AddScoped<IWordHandler, WordHandler>();
            services.AddScoped<IExcelHandler, ExcelHandler>();
        })
        .UseSerilog()
        .Build();

    return host;
}
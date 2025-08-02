using Konverter.Models;
using Konverter.Services;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.CommandLine;
using System.IO;

namespace Konverter
{
  public class Program
  {
    private static IHost ServiceHost { get; set; }

    public static T GetService<T>() where T : notnull => ServiceHost.Services.GetRequiredService<T>();

    public static async Task<int> Main(string[] args)
    {
      Option<FileInfo> excelFileOption = new("--excelFile")
      {
        Description = "Die Excel-Tabelle mit den Regie-Anweisungen",
        Required = true
      };

      Option<FileInfo[]> templatesOption = new("--template")
      {
        Description = "Eine PowerPoint-Vorlage, die für die Folien verwendet werden soll",
        Required = true,
        AllowMultipleArgumentsPerToken = true
      };

      RootCommand rootCommand = new("Tool zum Konvertieren einer Regie-Excel-Liste in PowerPoints");
      rootCommand.Options.Add(excelFileOption);
      rootCommand.Options.Add(templatesOption);

      rootCommand.SetAction(async parseResult => 
      {
        var excelFile = parseResult.GetValue(excelFileOption);
        var templates = parseResult.GetValue(templatesOption);

        await RunConverter(excelFile, templates);
      });

      ParseResult parseResult = rootCommand.Parse(args);
      return await parseResult.InvokeAsync();
    }

    private static async Task RunConverter(FileInfo excelFile, FileInfo[] powerPointTemplates)
    {
      Startup();

      var converterService = GetService<IConverterService>();

      await converterService.Convert(excelFile, powerPointTemplates);
    }

    private static void Startup()
    {
      var builder = Host.CreateDefaultBuilder();

      builder.ConfigureAppConfiguration((config) =>
      {
        config.AddJsonFile("appsettings.json");
      });

      builder.ConfigureServices((config, services) =>
      {
        services.Configure<DropboxConfig>(config.Configuration.GetSection("Dropbox"));
        services.Configure<ExcelConfig>(config.Configuration.GetSection("Excel"));
        services.Configure<ConverterConfig>(config.Configuration.GetSection("Converter"));
        services.Configure<OnedriveConfig>(config.Configuration.GetSection("OneDrive"));

        services.AddSingleton<IDropboxService, DropboxService>();
        services.AddSingleton<IOnedriveService, OnedriveService>();
        services.AddSingleton<IExcelService, ExcelService>();
        services.AddSingleton<IPowerpointService, PowerpointService>();
        services.AddSingleton<IConverterService, ConverterService>();
      });

      ServiceHost = builder.Build();
      ServiceHost.Start();
    }
  }
}

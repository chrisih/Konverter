using Konverter.Models;
using Konverter.Services;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Windows;

namespace Konverter
{
  public partial class App : Application
  {
    private static IHost ServiceHost { get; set; }

    public static T GetService<T>() where T : notnull => ServiceHost.Services.GetRequiredService<T>();

    public App()
    {
      InitializeComponent();

      var builder = Host.CreateDefaultBuilder();

      builder.ConfigureAppConfiguration((config) =>
      {
        config.AddJsonFile("appsettings.json");
      });

      builder.ConfigureServices((config, services) =>
      {
        services.Configure<DropboxConfig>(config.Configuration.GetSection("Dropbox"));
        services.Configure<ExcelConfig>(config.Configuration.GetSection("Excel"));
        services.Configure<PowerpointConfig>(config.Configuration.GetSection("Powerpoint"));
        services.Configure<ConverterConfig>(config.Configuration.GetSection("Converter"));

        services.AddSingleton<IDropboxService, DropboxService>();
        services.AddSingleton<IExcelService, ExcelService>();
        services.AddSingleton<IPowerpointService, PowerpointService>();
        services.AddSingleton<IConverterService, ConverterService>();
      });

      ServiceHost = builder.Build();
      ServiceHost.Start();
    }
  }
}
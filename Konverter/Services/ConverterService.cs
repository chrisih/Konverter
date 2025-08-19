using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Options;
using System.IO;


namespace Konverter.Services
{
  public class ConverterService : IConverterService
  {
    private readonly IExcelService _excelSvc;
    private readonly IPowerpointService _pptSvc;
    private readonly ConverterConfig _config;
    
    public ConverterService(IOptionsSnapshot<ConverterConfig> config, IOptionsSnapshot<PowerpointConfig> pptConfig, IExcelService excelSvc, IPowerpointService pptSvc)
    {
      _excelSvc = excelSvc;
      _pptSvc = pptSvc;
      _config = config.Value;
    }

    public async Task Convert(FileInfo excelfile, FileInfo template)
    {
      var slideTemplates = _excelSvc.GetSlideTemplates(excelfile).ToList();
      
      _pptSvc.CreatePresentationFromTemplate(template.FullName, $"{Path.GetTempFileName()}", slideTemplates);
    }
  }
}

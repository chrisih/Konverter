using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Options;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;


namespace Konverter.Services
{
  public class ConverterService : IConverterService
  {
    private readonly IExcelService _excelSvc;
    private readonly IPowerpointService _pptSvc;
    private readonly ConverterConfig _config;
    private readonly IOnedriveService _onedriveSvc;

    public ConverterService(IOptionsSnapshot<ConverterConfig> config, IOptionsSnapshot<PowerpointConfig> pptConfig, IExcelService excelSvc, IPowerpointService pptSvc, IOnedriveService onedriveSvc)
    {
      _excelSvc = excelSvc;
      _pptSvc = pptSvc;
      _config = config.Value;
      _onedriveSvc = onedriveSvc;
    }

    public async Task Convert(FileInfo excelfile, FileInfo template)
    {
      var slideTemplates = _excelSvc.GetSlideTemplates(excelfile).ToList();
      var presentation = _pptSvc.CreatePresentation(template);
      _pptSvc.AddSlides(presentation, slideTemplates);

      _onedriveSvc.Save(presentation);
    }
  }
}

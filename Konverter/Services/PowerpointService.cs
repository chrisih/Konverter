using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Extensions.Options;
using Konverter.Services.Abstraction;
using Konverter.Models;

namespace Konverter.Services
{
  public class PowerpointService : IPowerpointService
  {
    private IOptionsMonitor<PowerpointConfig> _config;

    public PowerpointService(IOptionsMonitor<PowerpointConfig> config)
    {
      _config = config;
    }

    public PowerPointApp CreatePowerpointApp()
    {
      return new PowerPointApp();
    }

    public Presentation CreatePresentation(PowerPointApp app)
    {
      return app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoCTrue);
    }

    public IEnumerable<PptTemplateConfig> GetPptTemplates() => _config.CurrentValue.PresentationTemplates;

    public void ApplyTemplate(Presentation presentation, string template)
    {
      presentation.ApplyTemplate(_config.CurrentValue.PresentationTemplates.Single(t => t.TemplateName == template).TemplatePath);
    }

    public void SetSize(PpSlideSizeType size, Presentation presentation)
    {
      presentation.PageSetup.SlideSize = size;
    }

    public IEnumerable<CustomLayout> GetCustomLayouts(Presentation presentation)
    {
      return presentation.SlideMaster.CustomLayouts.OfType<CustomLayout>();
    }
  }
}
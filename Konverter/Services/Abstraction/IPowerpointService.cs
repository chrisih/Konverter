using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using Microsoft.Office.Interop.PowerPoint;
using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IPowerpointService
  {
    PowerPointApp CreatePowerpointApp();
    Presentation CreatePresentation(PowerPointApp app);
    IEnumerable<PptTemplateConfig> GetPptTemplates();
    void ApplyTemplate(Presentation presentation, string template);
    void SetSize(PpSlideSizeType size, Presentation presentation);
    IEnumerable<CustomLayout> GetCustomLayouts(Presentation presentation);
  }
}
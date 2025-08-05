using System.IO;
using Konverter.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services.Abstraction
{
  public interface IPowerpointService
  {
    Presentation CreatePresentation();
    void SetSize(Presentation presentation, FileInfo file);
    IEnumerable<SlideTemplateFromExcel> TryGetPowerPointSlidesAsImage(SlideTemplateFromExcel template);
    IEnumerable<CustomLayout> GetCustomLayouts(Presentation presentation);
  }
}
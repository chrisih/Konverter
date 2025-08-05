using System.IO;
using Konverter.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services.Abstraction
{
  public interface IPowerpointService
  {
    Presentation CreatePresentation(FileInfo templateFile);
    void AddSlides(Presentation presentation, IEnumerable<SlideTemplateFromExcel> templates);
  }
}
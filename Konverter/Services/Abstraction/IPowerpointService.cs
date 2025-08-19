using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IPowerpointService
  {
    void CreatePresentationFromTemplate(string templatePath, string outputPath, IEnumerable<SlideTemplateFromExcel> partialProcessingInformation);
  }
}
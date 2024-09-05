using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services.Abstraction
{
  public interface IConverterService
  {
    IEnumerable<Presentation> Convert(string excelfile);
  }
}

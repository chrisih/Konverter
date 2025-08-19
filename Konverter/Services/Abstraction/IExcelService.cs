using System.IO;
using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IExcelService
  {
    IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(FileInfo excelFile);
  }
}

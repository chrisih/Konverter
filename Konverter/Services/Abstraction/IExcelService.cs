using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IExcelService
  {
    ExcelApp CreateExcelApp();
    Workbook OpenWorkbook(ExcelApp app, string path);
    Worksheet GetWorksheet(Workbook workbook, int index);
    IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(Worksheet contentSheet);
  }
}

using System.IO;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IExcelService
  {
    Workbook OpenWorkbook(ExcelApp app, FileInfo path);
    IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(FileInfo excelfile);
  }
}

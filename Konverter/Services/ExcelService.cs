using System.IO;
using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Konverter.Models;
using Microsoft.Extensions.Options;
using Konverter.Services.Abstraction;

namespace Konverter.Services
{
  public class ExcelService : IExcelService
  {
    private readonly IOptionsMonitor<ExcelConfig> _config;

    public ExcelService(IOptionsMonitor<ExcelConfig> config)
    {
      _config = config;
    }

    public ExcelApp CreateExcelApp()
    {
      return new ExcelApp();
    }

    public Workbook OpenWorkbook(ExcelApp app, FileInfo file)
    {
      return app.Workbooks.Open(file.FullName);
    }

    public Worksheet GetWorksheet(Workbook workbook, int index)
    {
      return workbook.Worksheets[index] as Worksheet;
    }

    public IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(Worksheet contentSheet)
    {
      for (int rowNum = 4; rowNum < 200; rowNum++)
      {
        var typeCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Type"]}{rowNum}"];
        if (typeCell.Value == null)
          continue;
        var contentCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Content"]}{rowNum}"];
        var titleCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Title"]}{rowNum}"];
        var footerCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Footer"]}{rowNum}"];
        var authorCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Author"]}{rowNum}"];
        var copyrightCell = contentSheet.Range[$"{_config.CurrentValue.Columns["Copyright"]}{rowNum}"];

        yield return new SlideTemplateFromExcel(typeCell, contentCell, titleCell, footerCell, authorCell, copyrightCell);
      }
    }
  }
}
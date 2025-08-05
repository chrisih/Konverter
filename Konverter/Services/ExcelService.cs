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
    private readonly ExcelApp _app;

    public ExcelService(IOptionsMonitor<ExcelConfig> config)
    {
      _config = config;
      _app = new ExcelApp();
    }
    
    public Workbook OpenWorkbook(ExcelApp app, FileInfo file)
    {
      return app.Workbooks.Open(file.FullName);
    }

    public Worksheet GetWorksheet(Workbook workbook, int index)
    {
      return workbook.Worksheets[index] as Worksheet;
    }

    public IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(FileInfo excelFile)
    {
      var workbook = _app.Workbooks.Open(excelFile.FullName);
      var schedule = workbook.Worksheets[2] as Worksheet;

      for (int rowNum = 4; rowNum < 200; rowNum++)
      {
        var typeCell = schedule.Range[$"{_config.CurrentValue.Columns["Type"]}{rowNum}"];
        if (typeCell.Value == null)
          continue;
        var contentCell = schedule.Range[$"{_config.CurrentValue.Columns["Content"]}{rowNum}"];
        var titleCell = schedule.Range[$"{_config.CurrentValue.Columns["Title"]}{rowNum}"];
        var footerCell = schedule.Range[$"{_config.CurrentValue.Columns["Footer"]}{rowNum}"];
        var authorCell = schedule.Range[$"{_config.CurrentValue.Columns["Author"]}{rowNum}"];
        var copyrightCell = schedule.Range[$"{_config.CurrentValue.Columns["Copyright"]}{rowNum}"];

        yield return new SlideTemplateFromExcel(typeCell, contentCell, titleCell, footerCell, authorCell, copyrightCell);
      }

      workbook.Close();
    }
  }
}
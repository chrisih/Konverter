using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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

    public IEnumerable<SlideTemplateFromExcel> GetSlideTemplates(FileInfo excelFile)
    {
      using var spreadsheetDocument = SpreadsheetDocument.Open(excelFile.FullName, false);

      var schedule = spreadsheetDocument.WorkbookPart?.WorksheetParts.ElementAtOrDefault(1)?.Worksheet;
      var strings = spreadsheetDocument.WorkbookPart?.SharedStringTablePart?.SharedStringTable;

      if (schedule == null)
        yield break;

      foreach(var row in schedule.Descendants<Row>())
      {
        var cells = row.Elements<Cell>().ToList();

        if(!cells.TryGetCellValue(_config.CurrentValue.Columns["Type"], strings, out var type) || string.IsNullOrWhiteSpace(type) || type == "Folientyp")
          continue;
        cells.TryGetCellValue(_config.CurrentValue.Columns["Content"], strings, out var content);
        cells.TryGetCellValue(_config.CurrentValue.Columns["Title"], strings, out var title);
        cells.TryGetCellValue(_config.CurrentValue.Columns["Footer"], strings, out var footer);
        cells.TryGetCellValue(_config.CurrentValue.Columns["Author"], strings, out var author);
        cells.TryGetCellValue(_config.CurrentValue.Columns["Copyright"], strings, out var copyright);

        yield return new SlideTemplateFromExcel
        {
          Content = content,
          Author = author,
          Copyright = copyright,
          Footer = footer,
          Title = title
        };
      }
    }
  }

  public static class CellHelpers
  {
    public static bool TryGetCellValue(this IEnumerable<Cell> cells, int index, SharedStringTable? strings, out string? value)
    {
      var cell = cells.ElementAtOrDefault(index);
      if (cell == null || string.IsNullOrWhiteSpace(cell.CellValue?.Text))
      {
        value = null;
        return false;
      }

      if (cell.DataType.Value == CellValues.SharedString)
      {
        value = strings.ElementAt(int.Parse(cell.InnerText)).InnerText;
        return true;
      }
      value = cell.CellValue.Text;
      return true;
    }
  }
}
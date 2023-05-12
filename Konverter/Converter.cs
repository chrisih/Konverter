using DevExpress.Mvvm;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Action = System.Action;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;

namespace Konverter;

public class Converter : ViewModelBase
{
  private readonly string _excelFile;
  private readonly string _templateFile;
  
  public Converter(string excelFile, string templateFile)
  {
    _excelFile = excelFile;
    _templateFile = templateFile;
  }

  public Presentation Convert(PpSlideSizeType? type = null)
  {
    Execute(() => Excel = new ExcelApp());
    Execute(() => PowerPoint = new PowerPointApp());

    Execute(() => Schedule = Excel.Workbooks.Open(_excelFile));
    Execute(() => Result = PowerPoint.Presentations.Add(MsoTriState.msoCTrue));
    Execute(() => Result.ApplyTemplate(_templateFile));
    if(type != null)
      Result.PageSetup.SlideSize = type.Value;
    Execute(() => ProcessSheet = Schedule.Worksheets[2] as Worksheet);
    Execute(() => CustomLayouts = GetLayouts().ToList());

    IterateAndCreateSlides();
    
    Schedule.Close(0);
    Excel.Quit();

    return Result;
  }

  private void IterateAndCreateSlides()
  {
    for (int rowNum = 4; rowNum < 200; rowNum++)
    {
      var typeCell = ProcessSheet.Range[$"B{rowNum}"];
      if (typeCell.Value == null)
        continue;
      var contentCell = ProcessSheet.Range[$"C{rowNum}"];
      var titleCell = ProcessSheet.Range[$"D{rowNum}"];
      var footerCell = ProcessSheet.Range[$"E{rowNum}"];
      var authorCell = ProcessSheet.Range[$"F{rowNum}"];
      var copyrightCell = ProcessSheet.Range[$"G{rowNum}"];
      var converter = new SingleRowConverter(typeCell, titleCell, contentCell, footerCell, authorCell, copyrightCell, CustomLayouts);
      converter.Convert(Result, Execute);
    }
  }

  private IEnumerable<CustomLayout> GetLayouts()
  {
    foreach (CustomLayout l in Result.SlideMaster.CustomLayouts)
      yield return l;
  }

  private void Execute(Action act)
  {
    Progress++;
    act();
    Progress++;
  }

  private List<CustomLayout> CustomLayouts { get; set; }
  private Presentation Result { get; set; }
  private Workbook Schedule { get; set; }
  private Worksheet ProcessSheet { get; set; }
  private PowerPointApp PowerPoint { get; set; }
  private ExcelApp Excel { get; set; }

  public int Progress
  {
    get => GetValue<int>();
    set => SetValue(value);
  }
}
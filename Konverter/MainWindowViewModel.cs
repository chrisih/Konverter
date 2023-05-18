using System.IO;
using System.Windows.Input;
using DevExpress.Mvvm;
using DevExpress.Mvvm.UI;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace Konverter
{
  public class MainWindowViewModel : ViewModelBase
  {
    public MainWindowViewModel()
    {
      BrowseExcelCommand = new DelegateCommand(BrowseExcel);
      BrowseStreamCommand = new DelegateCommand(BrowseStream);
      BrowseBeamerCommand = new DelegateCommand(BrowseBeamer);
      OpenExcelCommand = new DelegateCommand(OpenExcel);
      OpenStreamCommand = new DelegateCommand(OpenStream);
      OpenBeamerCommand = new DelegateCommand(OpenBeamer);
      CreateCommand = new AsyncCommand(Create, CanCreate);

      ExcelSheetFileName = Properties.Settings.Default.ExcelTemplate;
      StreamTemplateFileName = Properties.Settings.Default.StreamTemplate;
      BeamerTemplateFileName = Properties.Settings.Default.BeamerTemplate;
    }

    public ICommand OpenBeamerCommand { get; }

    private void OpenBeamer()
    {
      var app = new PowerPointApp();
      app.Presentations.Open(BeamerTemplateFileName);
    }

    public ICommand OpenStreamCommand { get; }

    private void OpenStream()
    {
      var app = new PowerPointApp();
      app.Presentations.Open(StreamTemplateFileName);
    }

    public ICommand OpenExcelCommand { get; }

    private void OpenExcel()
    {
      var app = new ExcelApp();
      var workbook = app.Workbooks.Open(ExcelSheetFileName);
      app.Visible = true;
      app.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
    }

    public ICommand BrowseExcelCommand { get; }
    private void BrowseExcel()
    {
      var svc = new OpenFileDialogService();
      svc.Multiselect = false;
      svc.CheckFileExists = true;
      svc.Filter = "Excel-Dateien|*.xls*";
      svc.ShowDialog();
      ExcelSheetFileName = svc.GetFullFileName();
    }

    public ICommand BrowseStreamCommand { get; }
    private void BrowseStream()
    {
      var svc = new OpenFileDialogService();
      svc.Multiselect = false;
      svc.CheckFileExists = true;
      svc.Filter = "PowerPoint-Vorlagen|*.potx";
      svc.ShowDialog();
      StreamTemplateFileName = svc.GetFullFileName();
    }

    public ICommand BrowseBeamerCommand { get; }
    private void BrowseBeamer()
    {
      var svc = new OpenFileDialogService();
      svc.Multiselect = false;
      svc.CheckFileExists = true;
      svc.Filter = "PowerPoint-Vorlagen|*.potx";
      svc.ShowDialog();
      BeamerTemplateFileName = svc.GetFullFileName();
    }

    public ICommand CreateCommand { get; }
    private bool CanCreate()
    {
      if (!File.Exists(ExcelSheetFileName))
        return false;
      if (!File.Exists(StreamTemplateFileName) && !File.Exists(BeamerTemplateFileName))
        return false;
      return true;
    }

    private async Task Create()
    {
      if (File.Exists(StreamTemplateFileName))
      {
        Converter = new Converter(ExcelSheetFileName, StreamTemplateFileName);
        Converter.Convert();
      }

      if (File.Exists(BeamerTemplateFileName))
      {
        Converter = new Converter(ExcelSheetFileName, BeamerTemplateFileName);
        Converter.Convert(PpSlideSizeType.ppSlideSizeOnScreen);
      }
    }

    public Converter Converter
    {
      get => GetValue<Converter>();
      set => SetValue(value);
    }

    public string ExcelSheetFileName
    {
      get => GetValue<string>();
      set => SetValue(value, SaveExcelFileName);
    }

    public string StreamTemplateFileName
    {
      get => GetValue<string>();
      set => SetValue(value, SaveStreamFileName);
    }

    public string BeamerTemplateFileName
    {
      get => GetValue<string>();
      set => SetValue(value, SaveBeamerFileName);
    }

    private void SaveExcelFileName()
    {
      Properties.Settings.Default.ExcelTemplate = ExcelSheetFileName;
      Properties.Settings.Default.Save();
    }

    private void SaveStreamFileName()
    {
      Properties.Settings.Default.StreamTemplate = StreamTemplateFileName;
      Properties.Settings.Default.Save();
    }

    private void SaveBeamerFileName()
    {
      Properties.Settings.Default.BeamerTemplate = BeamerTemplateFileName;
      Properties.Settings.Default.Save();
    }
  }
}

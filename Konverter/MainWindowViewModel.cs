using System.IO;
using System.Windows.Input;
using DevExpress.Mvvm;
using DevExpress.Mvvm.UI;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Konverter.Services.Abstraction;
using Konverter.Services;

namespace Konverter
{
  public class MainWindowViewModel : ViewModelBase
  {
    public MainWindowViewModel()
    {
      BrowseExcelCommand = new DelegateCommand(BrowseExcel);
      OpenExcelCommand = new DelegateCommand(OpenExcel);
      CreateCommand = new AsyncCommand(Create, CanCreate);

      ExcelSheetFileName = Properties.Settings.Default.ExcelTemplate;
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

    public ICommand CreateCommand { get; }
    private bool CanCreate()
    {
      if (!File.Exists(ExcelSheetFileName))
        return false;
      return true;
    }

    private async Task Create()
    {
      var onedrive = App.GetService<IOnedriveService>();
      await onedrive.Login();

      var converter = App.GetService<IConverterService>();

      await converter.Convert(ExcelSheetFileName);
    }

    public string ExcelSheetFileName
    {
      get => GetValue<string>();
      set => SetValue(value, SaveExcelFileName);
    }

    private void SaveExcelFileName()
    {
      Properties.Settings.Default.ExcelTemplate = ExcelSheetFileName;
      Properties.Settings.Default.Save();
    }
  }
}

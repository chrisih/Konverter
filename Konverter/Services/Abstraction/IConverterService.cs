namespace Konverter.Services.Abstraction
{
  public interface IConverterService
  {
    Task Convert(string excelfile);
  }
}

using System.IO;

namespace Konverter.Services.Abstraction
{
  public interface IConverterService
  {
    Task Convert(FileInfo excelfile, FileInfo template);
  }
}

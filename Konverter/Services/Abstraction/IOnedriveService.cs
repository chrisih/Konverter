using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services.Abstraction
{
  public interface IOnedriveService
  {
    Task Logon();
    Task Save(Presentation presentation);
  }

}

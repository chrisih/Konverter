using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Services
{
  public interface IOnedriveService
  {
    Task Save(Presentation presentation);
  }

  public class OnedriveService : IOnedriveService
  {
    public OnedriveService() 
    {

    }

    public async Task Save(Presentation presentation)
    {

    }
  }
}

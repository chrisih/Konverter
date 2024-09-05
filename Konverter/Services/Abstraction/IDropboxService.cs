using Dropbox.Api.Files;
using Konverter.Models;

namespace Konverter.Services.Abstraction
{
  public interface IDropboxService
  {
    IAsyncEnumerable<DropboxElement> GetFileInfoAsync(Predicate<FileMetadata>? filter = null);
    Task<byte[]> DownloadFileAsync(DropboxElement element);
  }
}
using Dropbox.Api;
using Dropbox.Api.Files;
using Konverter.Extensions;
using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Options;

namespace Konverter.Services
{
  public class DropboxService : IDropboxService
  {
    private readonly DropboxClient _client;
    private readonly DropboxConfig _settings;
    public DropboxService(IOptionsSnapshot<DropboxConfig> settings)
    {
      _client = new DropboxClient(settings.Value.Token);
      _settings = settings.Value;
    }

    public async Task<byte[]> DownloadFileAsync(DropboxElement element)
    {
      var link = await _client.Sharing.GetSharedLinkFileAsync(url: element.BaseLink.Url, path: element.Path);
      return await link.GetContentAsByteArrayAsync();
    }

    public async IAsyncEnumerable<DropboxElement> GetFileInfoAsync(Predicate<FileMetadata>? filter = null)
    {
      foreach (var url in _settings.Shares)
      {
        var sharedLink = new SharedLink(url);

        var baseElement = new DropboxElement(null, string.Empty, sharedLink);

        await foreach (var ret in _client.GetFolderContentRecursive(baseElement))
        {
          if (ret.RawData.IsFile && (filter == null || filter(ret.RawData as FileMetadata)))
            yield return ret;
        }
      }
    }
  }
}
using DevExpress.Mvvm.Native;
using Dropbox.Api;
using Dropbox.Api.Files;
using Konverter.Models;

namespace Konverter.Extensions
{
  public static class DropboxExtensions
  {
    public static async IAsyncEnumerable<DropboxElement> GetFolderContentRecursive(this DropboxClient client, DropboxElement parent)
    {
      var args = new ListFolderArg(parent.Path, false, true, false, false, true, null, parent.BaseLink, null, false);
      var elements = await client.Files.ListFolderAsync(args);

      foreach (var subfolder in elements.Entries.Where(e => e.IsFolder))
      {
        var folder = new DropboxElement(subfolder, $"{parent.Path}/{subfolder.Name}", parent.BaseLink);
        var children = client.GetFolderContentRecursive(folder);

        await foreach (var element in children)
          yield return element;
      }

      foreach (var file in elements.Entries.Where(e => e.IsFile))
        yield return new DropboxElement(file, $"{parent.Path}/{file.Name}", parent.BaseLink);
    }
  }
}
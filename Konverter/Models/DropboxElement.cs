using Dropbox.Api.Files;

namespace Konverter.Models
{
  public class DropboxElement
  {
    public DropboxElement() { }
    public DropboxElement(SharedLink baseLink) => BaseLink = baseLink;
    public DropboxElement(Metadata? rawData, string path, SharedLink baseLink) : this(baseLink)
    {
      RawData = rawData;
      Path = path;
    }

    public Metadata? RawData { get; }
    public string Path { get; }
    public SharedLink BaseLink { get; }
  }
}
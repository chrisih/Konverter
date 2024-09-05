namespace Konverter.Models
{
  public class DropboxConfig
  {
    public string Token { get; set; }

    public List<string> Shares { get; set; } = new List<string> 
    {
      "https://www.dropbox.com/sh/c4tuhbjz4p0npv4/AACURQzuxX8rj8RFRZSfSlzRa?dl=0"
    };
  }
}
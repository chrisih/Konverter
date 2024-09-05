namespace Konverter.Models
{
  internal class OnedriveConfig
  {
    public string ClientId { get; set; }
    public string ClientSecret { get; set; }
    public List<string> Scopes { get; set; }
  }
}

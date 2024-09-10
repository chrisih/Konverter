namespace Konverter.Models
{
  public class OnedriveConfig
  {
    public string ClientId { get; set; }
    public string ClientSecret { get; set; } 
    public string TenantId { get; set; } 

    public List<string> Scopes { get; set; }
  }
}

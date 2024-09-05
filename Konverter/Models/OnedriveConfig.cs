namespace Konverter.Models
{
  public class OnedriveConfig
  {
    public string ClientId { get; set; } = "f16e659d-3182-4f18-b51b-553ecaefc5ca";
    public string ClientSecret { get; set; } = "sil8Q~pEXy3D3S-pQL0Yz2.7B6WfYkyAZcX4pb_v";
    public string TenantId { get; set; } = "5b7bb46f-d979-4f15-939c-8ed3211bcefc";

    public List<string> Scopes { get; set; } = new List<string>
    {
      "https://graph.microsoft.com/.default"
    };
  }
}

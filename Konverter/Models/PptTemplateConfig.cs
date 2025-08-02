using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Models
{
  public class PptTemplateConfig
  {
    public string TemplateName { get; set; }
    public string TemplatePath { get; set; }

    public int SizeMode { get; set; }
  }
}
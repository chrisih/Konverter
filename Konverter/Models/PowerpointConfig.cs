using Microsoft.Office.Interop.PowerPoint;

namespace Konverter.Models
{
  public class PowerpointConfig
  {
    public List<PptTemplateConfig> PresentationTemplates { get; set; } = new List<PptTemplateConfig>
    {
      new PptTemplateConfig
      {
        TemplateName = "Beamer",
        TemplatePath = "Beamer_Master.potx",
        SizeMode = (int)PpSlideSizeType.ppSlideSizeOnScreen
      },
      new PptTemplateConfig
      {
        TemplateName = "Stream",
        TemplatePath = "Stream_Master.potx",
        SizeMode = (int)PpSlideSizeType.ppSlideSizeOnScreen16x9
      }
    };
  }

  public class PptTemplateConfig
  {
    public string TemplateName { get; set; }
    public string TemplatePath { get; set; }

    public int SizeMode { get; set; }
  }
}
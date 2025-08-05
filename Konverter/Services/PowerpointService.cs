using System.IO;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Extensions.Options;
using Konverter.Services.Abstraction;
using Konverter.Models;

namespace Konverter.Services
{
  public class PowerpointService : IPowerpointService
  {
    private readonly IOptionsMonitor<PowerpointConfig> _config;
    private readonly PowerPointApp _pptApp;

    public PowerpointService(IOptionsMonitor<PowerpointConfig> config)
    {
      _config = config;
      _pptApp = new PowerPointApp();
    }

    public Presentation CreatePresentation()
    {
      return _pptApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoCTrue);
    }

    public void SetSize(Presentation presentation, FileInfo file)
    {
      var template = _pptApp.Presentations.Open(file.FullName, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoFalse);

      var sixteennine = template.SlideMaster.Width / template.SlideMaster.Height == (float)16 / 9;

      presentation.PageSetup.SlideSize = sixteennine ? PpSlideSizeType.ppSlideSizeOnScreen16x9 : template.PageSetup.SlideSize;

      template.Close();
    }

    public IEnumerable<CustomLayout> GetCustomLayouts(Presentation presentation)
    {
      return presentation.SlideMaster.CustomLayouts.OfType<CustomLayout>();
    }

    public IEnumerable<SlideTemplateFromExcel> TryGetPowerPointSlidesAsImage(SlideTemplateFromExcel template)
    {
      if (string.IsNullOrWhiteSpace(template.Content) || !File.Exists(template.Content))
      {
        yield return template;
        yield break;
      }

      if (!template.Content.EndsWith(".ppt") || template.Content.EndsWith(".pptx"))
      {
        yield return template;
        yield break;
      }

      var toImport = _pptApp.Presentations.Open(template.Content, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoFalse);

      template = ReadFieldsFromPowerpoint(toImport, template);

      foreach (Slide sourceSlide in toImport.Slides)
      {
        yield return CreateImageSlideTemplate(sourceSlide, template);
      }

      toImport.Close();
    }

    private SlideTemplateFromExcel CreateImageSlideTemplate(Slide sourceSlide, SlideTemplateFromExcel template)
    {
      var tmpImagePath = Path.GetTempFileName() + ".png";
      sourceSlide.Export(tmpImagePath, "PNG", (int)sourceSlide.Master.Width * 2, (int)sourceSlide.Master.Height * 2);

      var toReturn = template;
      template.Content = tmpImagePath;

      return toReturn;
    }

    private SlideTemplateFromExcel ReadFieldsFromPowerpoint(Presentation toImport, SlideTemplateFromExcel template)
    {
      var ret = template;

      try
      {
        if (toImport.Slides[1].Shapes.Count >= 3)
        {
          TextFrame topmost = null;
          foreach (Shape shape in toImport.Slides[1].Shapes)
          {
            if (shape.TextFrame2.TextRange.Text.Contains("CCLI", StringComparison.InvariantCultureIgnoreCase))
            {
              ret.Copyright = shape.TextFrame2.TextRange.Text;
            }
            else if (shape.TextFrame2.TextRange.Text.Contains("Text", StringComparison.InvariantCultureIgnoreCase))
            {
              ret.Author = shape.TextFrame2.TextRange.Text;
            }
            else if (shape.TextFrame2.MarginTop < (topmost?.MarginTop ?? 1000))
            {
              topmost = (TextFrame)shape.TextFrame2;
              ret.Title = shape.TextFrame2.TextRange.Text;
            }
          }
        }
      }
      catch { }

      return ret;
    }
  }
}
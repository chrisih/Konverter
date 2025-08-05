using System.IO;
using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Options;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using TextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;

namespace Konverter.Services
{
  public class PowerpointService : IPowerpointService
  {
    private readonly IOptionsMonitor<PowerpointConfig> _config;
    private readonly PowerPointApp _pptApp;
    private readonly List<FieldConversionSetting> _settings;

    public PowerpointService(IOptionsMonitor<PowerpointConfig> config, IOptionsSnapshot<ConverterConfig> converterConfig)
    {
      _config = config;
      _settings = converterConfig.Value.FieldMappings;
      _pptApp = new PowerPointApp();
    }

    public void AddSlides(Presentation presentation, IEnumerable<SlideTemplateFromExcel> templates)
    {
      foreach (var template in templates)
      {
        var filledTemplates = FillTemplateAndExportSlides(template);

        foreach (var slideTemplate in filledTemplates)
        {
          var targetSlide = CreateTargetSlide(presentation, slideTemplate.PptLayoutReference);
          foreach (PowerPointShape shape in targetSlide.Shapes)
          {
            SetBasicShapeValues(presentation, shape, slideTemplate);
          }
        }
      }
    }

    private void SetBasicShapeValues(Presentation presentation, PowerPointShape shape, SlideTemplateFromExcel template)
    {
      var shapeName = GetShapeName(presentation, shape, template.PptLayoutReference);
      if (shapeName == null)
        return;

      var shapeType = _settings.FirstOrDefault(f => f.PowerpointFieldName == shapeName)?.InternalUsageType;

      switch (shapeType)
      {
        case FieldTypes.Titel:
          shape.TextFrame2.TextRange.Text = template.Title;
          shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
          break;
        case FieldTypes.Bild:
          if (!string.IsNullOrWhiteSpace(template.Content))
            shape.Fill.UserPicture(template.Content);
          break;
        case FieldTypes.Untertitel:
          shape.TextFrame2.TextRange.Text = template.Footer;
          shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
          break;
        case FieldTypes.Inhalt:
          shape.TextFrame2.TextRange.Text = template.Content;
          shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
          break;
        case FieldTypes.Autor:
          shape.TextFrame2.TextRange.Text = template.Author;
          shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
          break;
        case FieldTypes.Copyright:
          shape.TextFrame2.TextRange.Text = template.Copyright;
          shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
          break;
      }
    }

    private string? GetShapeName(Presentation presentation, PowerPointShape generatedShape, string? layoutName)
    {
      foreach (var shape in GetCustomLayouts(presentation, layoutName)?.Shapes?.OfType<PowerPointShape>())
        if (shape.Top == generatedShape.Top && shape.Left == generatedShape.Left && shape.Width == shape.Width && shape.Height == shape.Height)
          return shape.Name;
      return null;
    }

    private Slide CreateTargetSlide(Presentation presentation, string? layoutName)
    {
      var idx = presentation.Slides.Count + 1;
      var targetSlide = presentation.Slides.AddSlide(idx, GetCustomLayouts(presentation, layoutName));
      return targetSlide;
    }

    public Presentation CreatePresentation(FileInfo templateFile)
    {
      var presentation = _pptApp.Presentations.Add(MsoTriState.msoCTrue);

      presentation.ApplyTemplate(templateFile.FullName);
      
      var template = _pptApp.Presentations.Open(templateFile.FullName, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoFalse);

      var sixteennine = template.SlideMaster.Width / template.SlideMaster.Height == (float)16 / 9;

      presentation.PageSetup.SlideSize = sixteennine ? PpSlideSizeType.ppSlideSizeOnScreen16x9 : template.PageSetup.SlideSize;

      template.Close();

      return presentation;
    }

    private CustomLayout? GetCustomLayouts(Presentation presentation, string layoutName)
    {
      return presentation.SlideMaster.CustomLayouts.OfType<CustomLayout>().FirstOrDefault(l => l.Name == layoutName);
    }

    private IEnumerable<SlideTemplateFromExcel> FillTemplateAndExportSlides(SlideTemplateFromExcel template)
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

      var toImport = _pptApp.Presentations.Open(template.Content, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoFalse);

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
          foreach (PowerPointShape shape in toImport.Slides[1].Shapes)
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
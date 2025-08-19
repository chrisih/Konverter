using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Logging;
using RasterEdge.Imaging.Basic;
using RasterEdge.XDoc.PowerPoint;
using Picture = DocumentFormat.OpenXml.Presentation.Picture;
using Text = DocumentFormat.OpenXml.Presentation.Text;

public class OpenXmlPowerpointService : IPowerpointService
{
  private readonly ILogger<OpenXmlPowerpointService> _logger;

  public OpenXmlPowerpointService(ILogger<OpenXmlPowerpointService> logger)
  {
    _logger = logger;
  }

  public void CreatePresentationFromTemplate(string templatePath, string outputPath, IEnumerable<SlideTemplateFromExcel> partialProcessingInformation)
  {
    if (!File.Exists(templatePath))
      throw new FileNotFoundException($"Template file not found: {templatePath}");

    _logger.LogInformation("Creating '{file}'", outputPath);
    File.Copy(templatePath, outputPath, true);

    using var presentationFile = PresentationDocument.Open(outputPath, true);
    presentationFile.ChangeDocumentType(PresentationDocumentType.Presentation);
    var presentation = presentationFile.PresentationPart!;

    _logger.LogInformation("Extending processing list with exported linked powerpoints");
    var finalProcessingInformation = Extract(partialProcessingInformation);

    _logger.LogInformation("Creating powerpoint slides from processing list");
    var slides = CreateSlidesFromTemplates(presentation, finalProcessingInformation);

    _logger.LogInformation("Filling template shapes in created powerpoint");
    foreach (var slide in slides)
    {
      UpdateShapeContents(slide.Item1, slide.Item2);
    }

    // Save changes
    _logger.LogInformation("Saving changes");
    presentation.Presentation.Save();
  }

  private void UpdateShapeContents(SlidePart slide, SlideTemplateFromExcel processingInformation)
  {
    foreach (var shape in slide.Slide.Descendants<Text>())
    {
      Debugger.Break();
    }

    foreach (var shape in slide.Slide.Descendants<Picture>())
    {
      Debugger.Break();
    }
  }

  private IEnumerable<SlideTemplateFromExcel> Extract(IEnumerable<SlideTemplateFromExcel> partialProcessingInformation)
  {
    foreach (var processingInfo in partialProcessingInformation)
    {
      if (string.IsNullOrWhiteSpace(processingInfo.Content))
      {
        yield return processingInfo;
        continue;
      }


      if(!processingInfo.Content.EndsWith(".pptx", StringComparison.InvariantCultureIgnoreCase) && !processingInfo.Content.EndsWith(".ppt", StringComparison.InvariantCultureIgnoreCase))
      {
        yield return processingInfo;
        continue;
      }

      if (!File.Exists(processingInfo.Content))
      {
        yield return processingInfo;
        continue;
      }

      var updatedProcessingInfo = ReadFieldsFromPowerpoint(processingInfo.Content, processingInfo);

      var linkedDocument = new PPTDocument(processingInfo.Content);
      for (var pagenum = 0; pagenum < linkedDocument.GetPageCount(); pagenum++)
      {
        var page = (PPTXPage)linkedDocument.GetPage(pagenum);
        var filename = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
        page.ConvertToImage(ImageType.PNG, filename);
        updatedProcessingInfo.Content = filename;
        yield return updatedProcessingInfo;
      }

      linkedDocument.Dispose();
    }
  }

  private SlideTemplateFromExcel ReadFieldsFromPowerpoint(string pptFile, SlideTemplateFromExcel template)
  {
    var ret = template;

    using var toImport = PresentationDocument.Open(pptFile, true);
    foreach (var text in toImport.PresentationPart.SlideParts.First()
                                 .Slide.Descendants<Text>())
    {
      if (text.InnerText.Contains("ccli", StringComparison.InvariantCultureIgnoreCase))
        ret.Copyright = text.InnerText;
      else if (text.InnerText.Contains("Text", StringComparison.InvariantCultureIgnoreCase))
        ret.Author = text.InnerText;
      else
        ret.Title = text.InnerText;
    }

    return ret;
  }

  private IEnumerable<(SlidePart, SlideTemplateFromExcel)> CreateSlidesFromTemplates(PresentationPart presentation, IEnumerable<SlideTemplateFromExcel> templates)
  {
    var master = presentation.SlideMasterParts.First();

    foreach (var template in templates)
    {
      var layoutPart = master.SlideLayoutParts.FirstOrDefault(p => p.SlideLayout.CommonSlideData?.Name?.Value == template.PptLayoutReference);
      if (layoutPart != null)
      {
        var slide = CreateSlideFromLayout(presentation, layoutPart);
        
        yield return (slide, template);
      }
    }
  }

  private void ReplaceTextInSlide(SlidePart slidePart, Dictionary<string, string> placeholderValues)
  {
    var textElements = slidePart.Slide.Descendants<Text>();

    foreach (var textElement in textElements)
    {
      foreach (var placeholder in placeholderValues)
      {
        if (textElement.Text.Contains(placeholder.Key))
        {
          textElement.Text = textElement.Text.Replace(placeholder.Key, placeholder.Value);
        }
      }
    }
  }
  /*
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
  */
  private SlidePart CreateSlideFromLayout(PresentationPart presentationPart, SlideLayoutPart slideLayoutPart)
  {
    _logger.LogInformation("Creating slide...");

    // Create a new slide part
    var newSlidePart = presentationPart.AddNewPart<SlidePart>();

    // Create slide with basic structure
    newSlidePart.Slide = new Slide(
      new CommonSlideData(
        new ShapeTree(
          new NonVisualGroupShapeProperties(
            new NonVisualDrawingProperties { Id = 1U, Name = "" },
            new NonVisualGroupShapeDrawingProperties(),
            new ApplicationNonVisualDrawingProperties()),
          new GroupShapeProperties()
        )
      )
    );

    // Associate slide with layout
    newSlidePart.AddPart(slideLayoutPart);

    // Add to presentation slide list
    var slideIdList = presentationPart.Presentation.SlideIdList ??= new SlideIdList();

    var maxSlideId = slideIdList.Elements<SlideId>().Any() ? slideIdList.Elements<SlideId>().Max(x => x.Id!.Value) : 255U;

    var newSlideId = new SlideId
    {
      Id = maxSlideId + 1,
      RelationshipId = presentationPart.GetIdOfPart(newSlidePart)
    };

    slideIdList.Append(newSlideId);

    return newSlidePart;
  }
  
}
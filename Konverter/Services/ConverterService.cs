using Konverter.Models;
using Konverter.Services.Abstraction;
using Microsoft.Extensions.Options;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.IO;
using PowerPointShape = Microsoft.Office.Interop.PowerPoint.Shape;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace Konverter.Services
{
  public class ConverterService : IConverterService
  {
    private readonly IExcelService _excelSvc;
    private readonly IPowerpointService _pptSvc;
    private readonly ConverterConfig _config;
    private readonly PowerpointConfig _pptConfig;
    private readonly IOnedriveService _onedriveSvc;

    private ExcelApp _excelApp;

    public ConverterService(IOptionsSnapshot<ConverterConfig> config, IOptionsSnapshot<PowerpointConfig> pptConfig, IExcelService excelSvc, IPowerpointService pptSvc, IOnedriveService onedriveSvc)
    {
      _excelSvc = excelSvc;
      _pptSvc = pptSvc;
      _config = config.Value;
      _onedriveSvc = onedriveSvc;

      _pptConfig = pptConfig.Value;
      _excelApp = _excelSvc.CreateExcelApp();
    }

    private IEnumerable<CustomLayout> GetCustomLayouts(Presentation presentation) => presentation.SlideMaster.CustomLayouts.OfType<CustomLayout>();

    public async Task Convert(FileInfo excelfile, FileInfo template)
    {
      var workbook = _excelSvc.OpenWorkbook(_excelApp, excelfile);
      var schedule = _excelSvc.GetWorksheet(workbook, 2);
      var slideTemplates = _excelSvc.GetSlideTemplates(schedule);

      var presentation = _pptSvc.CreatePresentation();
      presentation.ApplyTemplate(template.FullName);
      _pptSvc.SetSize(presentation, template);

      foreach(var slideTemplate in slideTemplates)
      {
        AddSlide(presentation, slideTemplate);
      }

      _onedriveSvc.Save(presentation);
    }

    private void AddSlide(Presentation presentation, SlideTemplateFromExcel template)
    {
      var templates = _pptSvc.TryGetPowerPointSlidesAsImage(template);
      foreach (var slideTemplate in templates)
      {
        var targetSlide = CreateTargetSlide(presentation, slideTemplate.PptLayoutReference);
        foreach (PowerPointShape shape in targetSlide.Shapes)
        {
          SetBasicShapeValues(presentation, shape, slideTemplate);
        }
      }
    }

    private Slide CreateTargetSlide(Presentation presentation, string? layoutName)
    {
      var idx = presentation.Slides.Count + 1;
      var targetSlide = presentation.Slides.AddSlide(idx, GetCustomLayouts(presentation).FirstOrDefault(l => l.Name == layoutName));
      return targetSlide;
    }

    private string? GetShapeName(Presentation presentation, PowerPointShape generatedShape, string? layoutName)
    {
      foreach (var shape in GetCustomLayouts(presentation).FirstOrDefault(l => l.Name == layoutName)?.Shapes?.OfType<PowerPointShape>())
        if (shape.Top == generatedShape.Top && shape.Left == generatedShape.Left && shape.Width == shape.Width && shape.Height == shape.Height)
          return shape.Name;
      return null;
    }

    private void SetBasicShapeValues(Presentation presentation, PowerPointShape shape, SlideTemplateFromExcel template)
    {
      var shapeName = GetShapeName(presentation, shape, template.PptLayoutReference);
      if (shapeName == null)
        return;

      var shapeType = _config.FieldMappings.FirstOrDefault(f => f.PowerpointFieldName == shapeName)?.InternalUsageType;

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
  }
}

using System.IO;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Range = Microsoft.Office.Interop.Excel.Range;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Konverter;

public class SingleRowConverter
{
  private readonly CustomLayout _layout;
  private readonly string? _content;
  private readonly string? _title;
  private readonly string? _footer;
  private readonly string? _author;
  private readonly string? _copyright;

  public SingleRowConverter(Range typeCell, Range titleCell, Range contentCell, Range footerCell, Range authorCell, Range copyrightCell, List<CustomLayout> layouts)
  {
    _layout = layouts.Single(l => l.Name == typeCell.Value.ToString());
    _title = titleCell.Value?.ToString();
    _content = contentCell.Value?.ToString();
    _footer = footerCell.Value?.ToString();
    _author = authorCell.Value?.ToString();
    _copyright = copyrightCell.Value?.ToString();
  }

  private Presentation _presentation;
  private Action<Action> _iterator;

  public void Convert(Presentation p, Action<Action> iterator)
  {
    _presentation = p;
    _iterator = iterator;

    switch (_layout.Name)
    {
      case "Band-Lied":
        ImportPowerPoint();
        break;
      case "Predigt_mit_Folie":
        ImportPowerPoint();
        break;
      default:
        CreateSingleSlide();
        break;
    }
  }

  private void ImportPowerPoint()
  {
    _iterator(() => { });
    var toImport = _presentation.Application.Presentations.Open(_content, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoFalse);

    foreach (Slide sourceSlide in toImport.Slides)
    {
      _iterator(() => { });
      var tmpFile = Path.GetTempFileName();
      sourceSlide.Shapes.Range().Export(tmpFile, PpShapeFormat.ppShapeFormatPNG, 0, 0, PpExportMode.ppClipRelativeToSlide);

      _iterator(() => { });
      var idx = _presentation.Slides.Count + 1;
      var targetSlide = _presentation.Slides.AddSlide(idx, _layout);
      
      foreach (Shape shape in targetSlide.CustomLayout.Shapes)
      {
        if (shape.Name == "Bild")
          targetSlide.Shapes.AddPicture2(tmpFile, MsoTriState.msoFalse, MsoTriState.msoCTrue, shape.Left, shape.Top, shape.Width, shape.Height, MsoPictureCompress.msoPictureCompressTrue);
        else
          SetBasicShapeValues(shape);
      }

      targetSlide.CustomLayout = _layout;
    }

    toImport.Close();
  }

  private void SetBasicShapeValues(Shape shape)
  {
    if (shape.Name == "Titel")
    {
      shape.TextFrame2.TextRange.Text = _title;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shape.Name == "Untertitel")
    {
      shape.TextFrame2.TextRange.Text = _footer;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shape.Name == "Inhalt")
    {
      shape.TextFrame2.TextRange.Text = _content;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shape.Name == "Autor")
    {
      shape.TextFrame2.TextRange.Text = _author;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shape.Name == "Copyright")
    {
      shape.TextFrame2.TextRange.Text = _copyright;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
  }

  private void CreateSingleSlide()
  {
    _iterator(() => { });

    var idx = _presentation.Slides.Count + 1;
    var targetSlide = _presentation.Slides.AddSlide(idx, _layout);

    foreach (Shape shape in targetSlide.CustomLayout.Shapes)
    {
      SetBasicShapeValues(shape);
    }

    targetSlide.CustomLayout = _layout;

    _iterator(() => { });
  }
}
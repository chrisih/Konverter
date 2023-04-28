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
  private Presentation _presentation;
  private Action<Action> _iterator;

  public SingleRowConverter(Range typeCell, Range titleCell, Range contentCell, Range footerCell, Range authorCell, Range copyrightCell, List<CustomLayout> layouts)
  {
    _layout = layouts.Single(l => l.Name == typeCell.Value.ToString());
    _title = titleCell.Value?.ToString() ?? string.Empty;
    _content = contentCell.Value?.ToString() ?? string.Empty;
    _footer = footerCell.Value?.ToString() ?? string.Empty;
    _author = authorCell.Value?.ToString() ?? string.Empty;
    _copyright = copyrightCell.Value?.ToString() ?? string.Empty;
  }

  public void Convert(Presentation p, Action<Action> iterator)
  {
    _presentation = p;
    _iterator = iterator;

    switch (_layout.Name)
    {
      case Constants.Bandlied:
        ImportPowerPoint();
        break;
      case Constants.Bildpredigt:
        ImportPowerPoint();
        break;
      default:
        CreateSingleSlide();
        break;
    }
  }

  private void ImportPowerPoint()
  {
    if (string.IsNullOrWhiteSpace(_content))
    {
      var idx = _presentation.Slides.Count + 1;
      var targetSlide = _presentation.Slides.AddSlide(idx, _layout);
      return;
    }

    _iterator(() => { });
    var toImport = _presentation.Application.Presentations.Open(_content, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoFalse);

    foreach (Slide sourceSlide in toImport.Slides)
    {
      _iterator(() => { });
      sourceSlide.Copy();

      _iterator(() => { });
      var idx = _presentation.Slides.Count + 1;
      var targetSlide = _presentation.Slides.AddSlide(idx, _layout);

      var shapes = new List<Shape>(targetSlide.Shapes.OfType<Shape>());

      foreach (Shape shape in shapes)
      {
        SetBasicShapeValues(shape, targetSlide);
      }

      targetSlide.CustomLayout = _layout;
    }

    toImport.Close();
  }

  private string GetShapeName(Shape generatedShape)
  {
    foreach(Shape shape in _layout.Shapes)
      if(shape.Top == generatedShape.Top && shape.Left == generatedShape.Left && shape.Width == shape.Width && shape.Height == shape.Height)
        return shape.Name;
    return null;
  }

  private void SetBasicShapeValues(Shape shape, Slide slide)
  {
    var shapeName = GetShapeName(shape);
    if (shapeName == null)
      return;

    if (shapeName == $"{_layout.Name}_{Constants.Titel}")
    {
      shape.TextFrame2.TextRange.Text = _title;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{_layout.Name}_{Constants.Bild}")
    {
      var pastedShape = slide.Shapes.PasteSpecial(DataType: PpPasteDataType.ppPasteBitmap)[1];
      pastedShape.Left = shape.Left;
      pastedShape.Width = shape.Width;
      pastedShape.Height = shape.Height;
      pastedShape.Top = shape.Top;
    }
    else if (shapeName == $"{_layout.Name}_{Constants.Untertitel}")
    {
      shape.TextFrame2.TextRange.Text = _footer;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{_layout.Name}_{Constants.Inhalt}")
    {
      shape.TextFrame2.TextRange.Text = _content;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{_layout.Name}_{Constants.Autor}")
    {
      shape.TextFrame2.TextRange.Text = _author;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{_layout.Name}_{Constants.Copyright}")
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
    var shapes = new List<Shape>(targetSlide.Shapes.OfType<Shape>());

    foreach (Shape shape in shapes)
    {
      SetBasicShapeValues(shape, targetSlide);
    }

    targetSlide.CustomLayout = _layout;

    _iterator(() => { });
  }
}
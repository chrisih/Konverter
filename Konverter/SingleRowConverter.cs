using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Range = Microsoft.Office.Interop.Excel.Range;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Windows;
using System.IO;

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
  private readonly bool _pasteAsShape;

  public SingleRowConverter(Range typeCell, Range titleCell, Range contentCell, Range footerCell, Range authorCell, Range copyrightCell, List<CustomLayout> layouts, bool pasteAsShape)
  {
    _pasteAsShape = pasteAsShape;
    _layout = layouts.Single(l => l.Name == typeCell.Value.ToString());
    _title = titleCell.Value?.ToString() ?? string.Empty;
    _content = contentCell.Value?.ToString() ?? _title;
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
      case Constants.Bild:
        if (string.IsNullOrWhiteSpace(_content))
          return;
        if(_content!.EndsWith(".ppt") || _content!.EndsWith(".pptx"))
          ImportPowerPoint();
        else
          CreateSingleSlide();
        break;
      default:
        CreateSingleSlide();
        break;
    }
  }

  private Slide CreateTargetSlide()
  {
    var idx = _presentation.Slides.Count + 1;
    var targetSlide = _presentation.Slides.AddSlide(idx, _layout); // PpSlideLayout.ppLayoutBlank
    return targetSlide;
  }

  private void ImportPowerPoint()
  {
    // no file --> show dummy slide
    if (string.IsNullOrWhiteSpace(_content) || !File.Exists(_content))
    {
      var targetSlide = CreateTargetSlide();
      var shapes = new List<Shape>(targetSlide.Shapes.OfType<Shape>());
      foreach (Shape shape in shapes)
        SetBasicShapeValues(shape, targetSlide);
      return;
    }

    _iterator(() => { });

    // import possible
    if (_pasteAsShape)
    {
      _presentation.Slides.InsertFromFile(_content, _presentation.Slides.Count);
      return;
    }

    // import as image
    var toImport = _presentation.Application.Presentations.Open(_content, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoFalse);

    foreach (Slide sourceSlide in toImport.Slides)
    {
      _iterator(() => { });
      var targetSlide = CreateTargetSlide();

      _iterator(() => { });
      sourceSlide.Copy();

      var shapes = new List<Shape>(targetSlide.Shapes.OfType<Shape>());

      foreach (Shape shape in shapes)
      {
        SetBasicShapeValues(shape, targetSlide);
      }
    }

    toImport.Close();
  }

  private string? GetShapeName(Shape generatedShape)
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

    if (shapeName == $"{Constants.Titel}")
    {
      shape.TextFrame2.TextRange.Text = _title;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{Constants.Bild}")
    {
      if (!Clipboard.ContainsImage())
        return;
      var pastedShape = slide.Shapes.PasteSpecial(DataType: PpPasteDataType.ppPastePNG)[1];
      try
      {
        Clipboard.Clear();
      }
      catch { }
      pastedShape.Width = shape.Width;
      pastedShape.Height = shape.Height;
      pastedShape.Left = shape.Left;
      pastedShape.Top = shape.Top;
      shape.Delete();
    }
    else if (shapeName == $"{Constants.Untertitel}")
    {
      shape.TextFrame2.TextRange.Text = _footer;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{Constants.Inhalt}")
    {
      shape.TextFrame2.TextRange.Text = _content;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{Constants.Autor}")
    {
      shape.TextFrame2.TextRange.Text = _author;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
    else if (shapeName == $"{Constants.Copyright}")
    {
      shape.TextFrame2.TextRange.Text = _copyright;
      shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
    }
  }

  private void CreateSingleSlide()
  {
    _iterator(() => { });

    var targetSlide = CreateTargetSlide();
    var shapes = new List<Shape>(targetSlide.Shapes.Placeholders.OfType<Shape>());

    foreach (Shape shape in shapes)
    {
      SetBasicShapeValues(shape, targetSlide);
    }

    _iterator(() => { });
  }
}
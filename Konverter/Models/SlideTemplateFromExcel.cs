using Range = Microsoft.Office.Interop.Excel.Range;

namespace Konverter.Models
{
  public class SlideTemplateFromExcel
  {
    public SlideTemplateFromExcel(Range type, Range content, Range title, Range footer, Range author, Range copyright)
    {
      PptLayoutReference = type.Value?.ToString();
      Content = content.Value?.ToString();
      Title = title.Value?.ToString();
      Footer = footer.Value?.ToString();
      Author = author.Value?.ToString();
    }

    public string? PptLayoutReference { get; set; }
    public string? Content { get; set; }
    public string? Title { get; set; }
    public string? Footer { get; set; }
    public string? Author { get; set; }
    public string? Copyright { get; set; }
  }
}

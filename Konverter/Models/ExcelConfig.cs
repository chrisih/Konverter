namespace Konverter.Models
{
  public class ExcelConfig
  {
    public Dictionary<string, string> Columns { get; set; } = new Dictionary<string, string>
    {
      { "Type", "B" },
      { "Content", "C" },
      { "Title", "D" },
      { "Footer", "E" },
      { "Author", "F" },
      { "Copyright", "G" }
    };
  }
}

namespace Konverter.Models
{
  public class ExcelConfig
  {
    public Dictionary<string, int> Columns { get; set; } = new Dictionary<string, int>
    {
      { "Type", 0 },
      { "Content", 1 },
      { "Title", 2 },
      { "Footer", 3 },
      { "Author", 4 },
      { "Copyright", 5 }
    };
  }
}

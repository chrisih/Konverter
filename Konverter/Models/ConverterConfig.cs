namespace Konverter.Models
{
  public class ConverterConfig
  {
    public List<FieldConversionSetting> FieldMappings = new List<FieldConversionSetting>
    {
      new FieldConversionSetting
      {
        ExcelFieldName = "Titel",
        PowerpointFieldName = "Titel",
        InternalUsageType = FieldTypes.Titel
      },
      new FieldConversionSetting
      {
        ExcelFieldName = "Bild",
        PowerpointFieldName = "Bild",
        InternalUsageType = FieldTypes.Bild
      },
      new FieldConversionSetting
      {
        ExcelFieldName = "Untertitel",
        PowerpointFieldName = "Untertitel",
        InternalUsageType = FieldTypes.Untertitel
      },
      new FieldConversionSetting
      {
        ExcelFieldName = "Inhalt",
        PowerpointFieldName = "Inhalt",
        InternalUsageType = FieldTypes.Inhalt
      },
      new FieldConversionSetting
      {
        ExcelFieldName = "Autor",
        PowerpointFieldName = "Autor",
        InternalUsageType = FieldTypes.Autor
      },
      new FieldConversionSetting
      {
        ExcelFieldName = "Copyright",
        PowerpointFieldName = "Copyright",
        InternalUsageType = FieldTypes.Copyright
      }
    };
  }

  public class FieldConversionSetting
  {
    public string ExcelFieldName { get; set; }
    public string PowerpointFieldName { get; set; }
    public FieldTypes InternalUsageType { get; set; }
  }

  public enum FieldTypes
  {
    Titel,
    Bild,
    Untertitel,
    Inhalt,
    Autor,
    Copyright
  };
}
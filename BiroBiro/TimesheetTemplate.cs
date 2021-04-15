using System.IO;
using System.Text.Json;

namespace BiroBiro
{
    public class TimesheetTemplate
    {
        public string FileName { get; set; }
        public string CellMonthYear { get; set; }
        public int RowStartDates { get; set; }
        public TimesheetTemplateWorkShift WorkShift1 { get; set; }
        public TimesheetTemplateWorkShift WorkShift2 { get; set; }
        public TimesheetTemplateWorkShift WorkShift3 { get; set; }

        public string GetFullFileName() => $"{FileName}.xlsx";

        public void SaveToFile(string fileName)
        {
            JsonSerializerOptions options = new() { IncludeFields = true, WriteIndented = true };
            File.WriteAllText(fileName, JsonSerializer.Serialize(this, options));
        }

        public static TimesheetTemplate GetFromFile(string fileName)
            => JsonSerializer.Deserialize<TimesheetTemplate>(File.ReadAllText(fileName));
    }
}

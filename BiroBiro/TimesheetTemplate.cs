using System.Text.Json;
using System.IO;

namespace BiroBiro
{
    public class TimesheetTemplate
    {
        public string FileName { get; set; }
        public string CellMonthYear { get; set; }
        public int RowStartDates { get; set; }
        public string CollumnStart1 { get; set; }
        public string CollumnEnd1 { get; set; }
        public string CollumnStart2 { get; set; }
        public string CollumnEnd2 { get; set; }
        public string CollumnStart3 { get; set; }
        public string CollumnEnd3 { get; set; }
        public int HourStart1 { get; set; }
        public int MinuteStart1 { get; set; }
        public int HourEnd1 { get; set; }
        public int MinuteEnd1 { get; set; }
        public int HourStart2 { get; set; }
        public int MinuteStart2 { get; set; }
        public int HourEnd2 { get; set; }
        public int MinuteEnd2 { get; set; }
        public int HourStart3 { get; set; }
        public int MinuteStart3 { get; set; }
        public int HourEnd3 { get; set; }
        public int MinuteEnd3 { get; set; }

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

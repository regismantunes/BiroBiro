using System;
using System.IO;

namespace BiroBiro
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            Console.WriteLine("I am Biro Biro and I will create and fill in the timesheet of this month for you!");

            try
            {
                DateTime startDate = args != null && args.Length > 0 ? 
                    DateTime.Parse(args[0]) : 
                    DateTime.Today.AddDays((DateTime.Today.Day - 1) * -1);

                Timesheet ts = new();
                const string defaultTemplateFileName = "template.json";
                if (File.Exists(defaultTemplateFileName))
                    ts.Template = TimesheetTemplate.GetFromFile(defaultTemplateFileName);
                else
#if RELEASE
                    throw new FileNotFoundException($"The file {defaultTemplateFileName} was not found.");
#endif
#if DEBUG
                {
                    ts.Template = new TimesheetTemplate()
                    {
                        FileName = "Planilha de Horas",
                        RowStartDates = 14,
                        CellMonthYear = "A7",
                        CollumnStart1 = "C",
                        CollumnEnd1 = "D",
                        CollumnStart2 = "E",
                        CollumnEnd2 = "F",
                        HourStart1 = 8,
                        HourEnd1 = 12,
                        HourStart2 = 13,
                        HourEnd2 = 17
                    };
                    ts.Template.SaveToFile(defaultTemplateFileName);
                }
#endif
                ts.CreateAndFillNewFile(startDate.Year, startDate.Month, startDate.Day);

                Console.WriteLine("The timesheet was created and completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.ReadKey();
            }
        }
    }
}

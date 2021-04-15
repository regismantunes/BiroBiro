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
                        WorkShift1 = new TimesheetTemplateWorkShift()
                        { 
                            CollumnStart = "C",
                            CollumnEnd = "D",
                            HourStart = 8,
                            HourEnd = 12,
                        },
                        WorkShift2 = new TimesheetTemplateWorkShift()
                        {
                            CollumnStart = "E",
                            CollumnEnd = "F",
                            HourStart = 13,
                            HourEnd = 17,
                        }
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

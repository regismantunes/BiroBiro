using System;

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
                ts.CreateAndFillNewFile(startDate.Year, startDate.Month, startDate.Day, true);

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

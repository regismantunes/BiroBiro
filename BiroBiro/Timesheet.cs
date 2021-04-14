using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace BiroBiro
{
    public class Timesheet
    {
        public string TemplateName { get; set; } = "Planilha de Horas";

        public string GetTemplateFileName() => $"{TemplateName}.xlsx";
        private string GetDefaultFileName(int year, int month) => $"{TemplateName} - {year:0000}-{month:00}.xlsx";
        private string GetDefaultFileNameN(int year, int month, int count) => $"{TemplateName} - {year:0000}-{month:00} ({count}).xlsx";

        private string CreateNewFile(int year, int month)
        {
            string templateFileName = GetTemplateFileName();
            if (!File.Exists(templateFileName))
                throw new FileNotFoundException($"The template file {templateFileName} was not found.");

            string newFileName = GetDefaultFileName(year, month);
            int count = 0;
            while (File.Exists(newFileName))
            {
                count++;
                newFileName = GetDefaultFileNameN(year, month, count);
            }

            File.Copy(templateFileName, newFileName);

            return newFileName;
        }

        private void FillFile(string fileName, int year, int month, int startDay = 1)
        {
            if (!File.Exists(fileName))
                throw new FileNotFoundException($"The file {fileName} was not found.");

            IReadOnlyList<DateTime> lstHolidays = Holidays.GetHolidays(year);

            Application excel = new();
            Workbooks wbs = excel.Workbooks;
            Workbook wb = wbs.Open(Path.Combine(AppContext.BaseDirectory, fileName), 0, false);
            try
            {
                Worksheet ws = wb.ActiveSheet;

                DateTime date = new(year, month, 1);
                ws.get_Range("A7").Value = date;

                Random rd = new();
                int row = 14;
                do
                {
                    if (date.Day >= startDay && 
                        date.DayOfWeek != DayOfWeek.Sunday && 
                        date.DayOfWeek != DayOfWeek.Saturday &&
                        !lstHolidays.Contains(date))
                    {
                        int dif = rd.Next(-15, 15);
                        ws.get_Range($"C{row}").Value = date.AddHours(8).AddMinutes(dif).ToString("HH:mm");
                        dif = rd.Next(-15, 15);
                        ws.get_Range($"D{row}").Value = date.AddHours(12).AddMinutes(dif).ToString("HH:mm");
                        dif = rd.Next(-15, 15);
                        ws.get_Range($"E{row}").Value = date.AddHours(13).AddMinutes(dif).ToString("HH:mm");
                        dif = rd.Next(-15, 15);
                        ws.get_Range($"F{row}").Value = date.AddHours(17).AddMinutes(dif).ToString("HH:mm");
                    }
                    row++;
                    date = date.AddDays(1);
                } while (date.Month == month);

                wb.Save();
            }
            finally
            {
                int hWnd = excel.Application.Hwnd;
                
                wb.Close();
                wbs.Close();
                excel.Quit();

                GetWindowThreadProcessId((IntPtr)hWnd, out uint processID);
                Process.GetProcessById((int)processID).Kill();
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public void CreateAndFillNewFile(int year, int month, int startDay = 1)
        {
            string fileName = CreateNewFile(year, month);
            FillFile(fileName, year, month, startDay);
        }
    }
}

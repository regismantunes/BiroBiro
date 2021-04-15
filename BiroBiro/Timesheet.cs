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
        public TimesheetTemplate Template { get; set; }

        private string GetDefaultFileName(int year, int month) => $"{Template.FileName} - {year:0000}-{month:00}.xlsx";
        private string GetDefaultFileNameN(int year, int month, int count) => $"{Template.FileName} - {year:0000}-{month:00}({count}).xlsx";

        private string CreateNewFile(int year, int month)
        {
            string templateFileName = Template.GetFullFileName();
            if (!File.Exists(templateFileName))
                throw new FileNotFoundException($"The template file {templateFileName} was not found.");

            int count = 0;
            string newFileName = GetDefaultFileName(year, month);
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
            Workbook wb = null;
            try
            {
                wb = wbs.Open(Path.Combine(AppContext.BaseDirectory, fileName), 0, false);
                Worksheet ws = wb.ActiveSheet;

                DateTime date = new(year, month, 1);
                if (!string.IsNullOrEmpty(Template.CellMonthYear))
                    ws.get_Range(Template.CellMonthYear).Value = date;

                Random rd = new();
                int row = Template.RowStartDates;
                do
                {
                    if (date.Day >= startDay &&
                        date.DayOfWeek != DayOfWeek.Sunday &&
                        date.DayOfWeek != DayOfWeek.Saturday &&
                        !lstHolidays.Contains(date))
                    {
                        int dif;
                        //ws.get_Range($"C{row}").Value = date.AddHours(8).AddMinutes(dif).ToString("HH:mm");
                        if (!string.IsNullOrEmpty(Template.CollumnStart1) &&
                            !string.IsNullOrEmpty(Template.CollumnEnd1) &&
                            Template.HourStart1 >= 0 && Template.HourStart1 < 24 &&
                            Template.MinuteStart1 >= 0 && Template.MinuteStart1 < 60 &&
                            Template.HourEnd1 >= 0 && Template.HourEnd1 < 24 &&
                            Template.MinuteEnd1 >= 0 && Template.MinuteEnd1 < 60)
                        {
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnStart1}{row}").Value = date.AddHours(Template.HourStart1).AddMinutes(Template.MinuteStart1 + dif).ToString("HH:mm");
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnEnd1}{row}").Value = date.AddHours(Template.HourEnd1).AddMinutes(Template.MinuteEnd1 + dif).ToString("HH:mm");
                        }
                        if (!string.IsNullOrEmpty(Template.CollumnStart2) &&
                            !string.IsNullOrEmpty(Template.CollumnEnd2) &&
                            Template.HourStart2 >= 0 && Template.HourStart2 < 24 &&
                            Template.MinuteStart2 >= 0 && Template.MinuteStart2 < 60 &&
                            Template.HourEnd2 >= 0 && Template.HourEnd2 < 24 &&
                            Template.MinuteEnd2 >= 0 && Template.MinuteEnd2 < 60)
                        {
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnStart2}{row}").Value = date.AddHours(Template.HourStart2).AddMinutes(Template.MinuteStart2 + dif).ToString("HH:mm");
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnEnd2}{row}").Value = date.AddHours(Template.HourEnd2).AddMinutes(Template.MinuteEnd2 + dif).ToString("HH:mm");
                        }
                        if (!string.IsNullOrEmpty(Template.CollumnStart3) &&
                            !string.IsNullOrEmpty(Template.CollumnEnd3) &&
                            Template.HourStart3 >= 0 && Template.HourStart3 < 24 &&
                            Template.MinuteStart3 >= 0 && Template.MinuteStart3 < 60 &&
                            Template.HourEnd3 >= 0 && Template.HourEnd3 < 24 &&
                            Template.MinuteEnd3 >= 0 && Template.MinuteEnd3 < 60)
                        {
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnStart3}{row}").Value = date.AddHours(Template.HourStart3).AddMinutes(Template.MinuteStart3 + dif).ToString("HH:mm");
                            dif = rd.Next(-15, 15);
                            ws.get_Range($"{Template.CollumnEnd3}{row}").Value = date.AddHours(Template.HourEnd3).AddMinutes(Template.MinuteEnd3 + dif).ToString("HH:mm");
                        }
                    }
                    row++;
                    date = date.AddDays(1);
                } while (date.Month == month);

                wb.Save();
            }
            finally
            {
                int hWnd = excel.Application.Hwnd;
                
                wb?.Close();
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

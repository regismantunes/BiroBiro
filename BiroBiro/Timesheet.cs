using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

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
                        if (Template.WorkShift1 != null)
                            FillCellsStartEnd(date, rd, ws, row, Template.WorkShift1);
                        if (Template.WorkShift2 != null)
                            FillCellsStartEnd(date, rd, ws, row, Template.WorkShift2);
                        if (Template.WorkShift3 != null)
                            FillCellsStartEnd(date, rd, ws, row, Template.WorkShift3);
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

        private void FillCellsStartEnd(DateTime date, Random rd, Worksheet ws, int row, TimesheetTemplateWorkShift workShift)
        {
            if (!string.IsNullOrEmpty(workShift.CollumnStart) &&
                !string.IsNullOrEmpty(workShift.CollumnEnd) &&
                workShift.HourStart >= 0 && workShift.HourStart < 24 &&
                workShift.MinuteStart >= 0 && workShift.MinuteStart < 60 &&
                workShift.HourEnd >= 0 && workShift.HourEnd < 24 &&
                workShift.MinuteEnd >= 0 && workShift.MinuteEnd < 60)
            {
                int dif = rd.Next(-15, 15);
                ws.get_Range($"{workShift.CollumnStart}{row}").Value = date.AddHours(workShift.HourStart).AddMinutes(workShift.MinuteStart + dif).ToString("HH:mm");
                dif = rd.Next(-15, 15);
                ws.get_Range($"{workShift.CollumnEnd}{row}").Value = date.AddHours(workShift.HourEnd).AddMinutes(workShift.MinuteEnd + dif).ToString("HH:mm");
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

using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;

namespace BiroBiro
{
    public class Timesheet
    {
        public string TemplateName { get; set; } = "Planilha de Horas";

        public string GetTemplateFileName() => $"{TemplateName}.xlsx";
        public string GetFileName(int year, int month) => $"{TemplateName} - {year:0000}-{month:00}.xlsx";
        
        private void CreateNewFile(int year, int month, bool overwrite = false)
        {
            string strTemplateFileName = GetTemplateFileName();
            if (!File.Exists(strTemplateFileName))
                throw new FileNotFoundException($"The template file {strTemplateFileName} was not found.");

            string strNomeNovoArquivo = GetFileName(year, month);
            if (File.Exists(strNomeNovoArquivo))
            {
                if (!overwrite)
                    throw new Exception($"The file {strNomeNovoArquivo} already exists.");

                File.Delete(strNomeNovoArquivo);
            }

            File.Copy(strTemplateFileName, strNomeNovoArquivo);
        }

        private void FillFile(int year, int month, int startDay = 1)
        {
            string strNomeNovoArquivo = GetFileName(year, month);
            if (!File.Exists(strNomeNovoArquivo))
                throw new FileNotFoundException($"The file {strNomeNovoArquivo} was not found.");

            IReadOnlyList<DateTime> lstHolidays = Holidays.GetHolidays(year);

            Application excel = new();
            Workbook wb = excel.Workbooks.Open(Path.Combine(AppContext.BaseDirectory, strNomeNovoArquivo), 0, false);
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
                wb.Close();
                excel.Quit();
            }
        }

        public void CreateAndFillNewFile(int year, int month, int startDay = 1, bool overwrite = false)
        {
            CreateNewFile(year, month, overwrite);
            FillFile(year, month, startDay);
        }
    }
}

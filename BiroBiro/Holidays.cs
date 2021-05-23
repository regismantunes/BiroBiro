using System;
using System.Collections.Generic;

namespace BiroBiro
{
    public static class Holidays
    {
        private static DateTime EasterDay(int year)
        {
            int x = 24;
            int y = 5;
            int a = year % 19;
            int b = year % 4;
            int c = year % 7;
            int d = ((19 * a) + x) % 30;
            int e = ((2 * b) + (4 * c) + (6 * d) + y) % 7;
            
            int day, month;
            if (d + e > 9)
            {
                day = d + e - 9;
                month = 4;
            }
            else
            {
                day = d + e + 22;
                month = 3;
            }
            
            if (month == 4)
            {
                if (day == 26)
                    day = 19;
                else if (day == 25 && d == 28 && a > 10)
                    day = 18;
            }

            return new DateTime(year, month, day);
        }

        public static IReadOnlyList<DateTime> GetHolidays(int y)
        {
            DateTime anoNovo = new(y, 1, 1);
            DateTime pascoa = EasterDay(y);
            DateTime carnaval1 = pascoa.AddDays(-48);
            DateTime carnaval2 = pascoa.AddDays(-47);
            DateTime paixaoCristo = pascoa.AddDays(-2);
            DateTime tiradentes = new(y, 4, 21);
            DateTime corpusChristi = pascoa.AddDays(60);
            DateTime diaTrabalho = new(y, 5, 1);
            DateTime diaIndependencia = new(y, 9, 7);
            DateTime nossaSenhora = new(y, 10, 12);
            DateTime finados = new(y, 11, 2);
            DateTime proclamaRepublica = new(y, 11, 15);
            DateTime natal = new(y, 12, 25);

            DateTime[] dates = { anoNovo, carnaval1, carnaval2, paixaoCristo, pascoa, tiradentes, corpusChristi, diaTrabalho, diaIndependencia, nossaSenhora, finados, proclamaRepublica, natal };
            return new List<DateTime>(dates);
        }
    }
}

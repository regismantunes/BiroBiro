using System;
using System.Collections.Generic;

namespace BiroBiro
{
    public static class Holidays
    {
        private static DateTime EasterDay(int y)
        {
            double c = Math.Floor((double)(y / 100));
            double n = y - 19 * Math.Floor((double)(y / 19));
            double k = Math.Floor((c - 17) / 25);
            double i = c - Math.Floor(c / 4) - Math.Floor((c - k) / 3) + 19 * n + 15;
            i -= 30 * Math.Floor((i / 30));
            i -= Math.Floor(i / 28) * (1 - Math.Floor(i / 28) * Math.Floor(29 / (i + 1)) * Math.Floor((21 - n) / 11));
            double j = y + Math.Floor((double)(y / 4)) + i + 2 - c + Math.Floor(c / 4);
            j -= 7 * Math.Floor(j / 7);
            double l = i - j;
            double m = 3 + Math.Floor((l + 40) / 44);
            double d = l + 28 - 31 * Math.Floor(m / 4);
            return new DateTime(y, (int)(m - 1), (int)d);
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

        /*private const string FILENAME = "holidays{0:0000}.json";

        private static void CreateFileHolidays(int year)
        {
            WebRequest request = WebRequest.Create($"https://holidayapi.com/v1/holidays?pretty&key=0e53b7ed-2276-49b1-a1a7-d9e2d5beb19f&country=BR&year={year:0000}");
            request.Credentials = CredentialCache.DefaultCredentials;
            ((HttpWebRequest)request).UserAgent = null;

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            if (response.StatusCode != HttpStatusCode.OK)
                throw new Exception(response.StatusCode.ToString());

            Stream receiveStream = response.GetResponseStream();

            FileStream fs = File.OpenWrite(string.Format(FILENAME, year));

            Span<byte> buffer = new();
            receiveStream.Read(buffer);
            fs.Write(buffer);

            fs.Close();
        }

        public static IReadOnlyList<DateTime> GetHolidays(int year)
        {
            string fileName = string.Format(FILENAME, year);
            
            if (!File.Exists(fileName))
                CreateFileHolidays(year);

            List<DateTime> lstDates = new();
            JsonDocument document = JsonDocument.Parse(File.ReadAllText(fileName));
            JsonElement jeHolidays = document.RootElement.GetProperty("holidays");
            
            lstDates.AddRange(from JsonElement jeHoliday in jeHolidays.EnumerateArray()
                              select jeHoliday.GetProperty("date").GetDateTime());

            return lstDates;
        }*/
    }
}

using Nager.Date;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grafikon.Model
{
    public class Satnica
    {
        public string ime { get; set; }

        public string prezime { get; set; }

        public int godina { get; set; }

        public int mjesec { get; set; }

        public int startWork { get; set; }

        public int endWork { get; set; }


        // return last day of month
        public DateTime LastDay()
        {
            return new DateTime(godina, mjesec, this.DaysInMonth());
        }

        // return first day of month
        public DateTime FirstDay()
        {
            return new DateTime(godina, mjesec, 1);
        }

        // Count total hours of work day
        public int TotalWork()
        {
            return this.endWork - this.startWork;
        }

        // Count num of days 
        public int DaysInMonth()
        {
            DateTime datum = new DateTime(this.godina, this.mjesec, 1);
            int count = 0;
            while (datum.Month == mjesec)
            {
                count++;
                datum = datum.AddDays(1);
            }

            return count;
        }

        // Open templet and give it to our workbook [initialize this method to variable Type HSSFWorkbook]
        public static HSSFWorkbook openTemp()
        {
            HSSFWorkbook workbook;
            using (FileStream file = new FileStream(Environment.CurrentDirectory + "\\templetGrafikon.xls", FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(file);
                file.Close();
            }
            return workbook;
        }

        // Save our edited workbook [Pass variable Type HSSFWorkbook which you create with openTemp method]
        public static void saveTemp(HSSFWorkbook workbook)
        {
            using (FileStream file = new FileStream(Environment.CurrentDirectory + "\\templetGrafikon2.xls", FileMode.CreateNew, FileAccess.Write))
            {
                workbook.Write(file);
                file.Close();
            }
        }


        // Check is given date holiday
        public bool holidayCheck(DateTime datum)
        {
            var isPublicHolday = DateSystem.IsPublicHoliday(datum, CountryCode.HR);
            return isPublicHolday;
            
        }
    }
}

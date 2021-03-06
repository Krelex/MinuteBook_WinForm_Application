﻿using Nager.Date;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
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
        public string nazivPoduzeca { get; set; }

        public string ime { get; set; }

        public string prezime { get; set; }

        public int godina { get; set; }

        public int mjesec { get; set; }

        public int startWork { get; set; }

        public int endWork { get; set; }

        public bool puerperal { get; set; }

        public bool FieldWork { get; set; }

        public bool vacation { get; set; }

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
        public static void saveTemp(HSSFWorkbook workbook , string saveFile , string savePath)
        {
            
            using (FileStream file = new FileStream(savePath + saveFile, FileMode.CreateNew, FileAccess.Write))
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


        // Create string for FileName from this object ime , mjesec and godina
        public string FileNameCreator ()
        {
            return this.ime.ToUpper() + "-" + this.mjesec + "-" + this.godina + ".xls";
        }

        // Set company prop value in corisponding sheet given as a parametar
        public void SetCompanyName(ISheet sheet)
        {
            sheet.GetRow(4).Cells[1].SetCellValue(this.nazivPoduzeca);
        }

        // Set name and surname prop value in corisponding sheet given as a parametar
        public void SetNameSurname ( ISheet sheet)
        {
            sheet.GetRow(6).Cells[1].SetCellValue(this.ime + " " + this.prezime);
        }

        // Set FirstDay method value in corisponding sheet given as a parametar
        public void SetFirstDayOfMonth(ISheet sheet)
        {
            sheet.GetRow(8).Cells[1].SetCellValue(this.FirstDay());
        }

        // Set FirstDay method value in corisponding sheet given as a parametar
        public void SetLastDayOfMonth(ISheet sheet)
        {
            sheet.GetRow(8).Cells[4].SetCellValue(this.LastDay());
        }

        // Populate Date and Day.string("ddd") Column in corisponding sheet given as a parametar
        public void SetDateAndDay(ISheet sheet, int row, DateTime datum)
        {
            sheet.GetRow(row).Cells[0].SetCellValue(datum.Date);
            sheet.GetRow(row).Cells[1].SetCellValue(datum.Date.ToString("ddd"));
        }

        // Populate StartWork Column in corisponding sheet given as a parametar
        public void SetStartWork(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[2].SetCellValue(this.startWork);
        }

        // Populate EndWork Column in corisponding sheet given as a parametar
        public void SetEndWork(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[3].SetCellValue(this.endWork);
        }

        // Populate TotalWork Column in corisponding sheet given as a parametar
        public void SetTotalWork(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[5].SetCellValue(this.TotalWork());
        }

        // Populate TotalPuerperal Column in corisponding sheet given as a parametar
        public void SetTotalPuerperal(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[15].SetCellValue(this.TotalWork());
        }

        // Populate TotalFieldWork Column in corisponding sheet given as a parametar
        public void SetFieldWork(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[11].SetCellValue(this.TotalWork());
        }

        // Populate TotalVocation Column in corisponding sheet given as a parametar
        public void SetTotalVacation(ISheet sheet, int row)
        {
            sheet.GetRow(row).Cells[13].SetCellValue(this.TotalWork());
        }

        // Method for open file from save location
        public void openFile(string path)
        {
            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbooks books = excel.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook sheet2 = books.Open(path);
        }
    }
}

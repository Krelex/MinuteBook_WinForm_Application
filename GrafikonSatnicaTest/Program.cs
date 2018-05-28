using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Grafikon.Model;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace GrafikonSatnicaTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string name = "fico";
            string pathh = @"C:\Users\" + Environment.UserName + @"\desktop\transaction\\"+name+".xsl";
          
            Console.WriteLine(pathh);
            //Create working object
            Satnica ob1 = new Satnica();

            //Open xls. temp and initialize it to HSSFWorkbook
            HSSFWorkbook workbook = Satnica.openTemp();

            //Point to sheet of workbook
            ISheet sheet = workbook.GetSheetAt(0);

            //Dohvacanje reda i kolone exemple
            //sheet.GetRow(11).Cells[0].SetCellValue("8");


            //Hardcode data
            ob1.godina = 2018;
            ob1.mjesec = 12;
            ob1.startWork = 9;
            ob1.endWork = 17;
            ob1.ime = "Sanda";
            ob1.prezime = "Nesto";
            ob1.puerperal = false;


            //Dynamic data
            //Console.WriteLine("Upisite IME zaposlenika");
            //ob1.ime = Console.ReadLine();
            //Console.WriteLine("Upisite PREZIME zaposlenika");
            //ob1.prezime = Console.ReadLine();
            //Console.WriteLine("Upisite GODINU format['yyyy']");
            //ob1.godina = int.Parse(Console.ReadLine());
            //Console.WriteLine("Upisite MJESEC format['MM']");
            //ob1.mjesec = int.Parse(Console.ReadLine());
            //Console.WriteLine("Upisite POCETAK RADA zaposlenika format['hh']");
            //ob1.startWork = int.Parse(Console.ReadLine());
            //Console.WriteLine("Upisite KRAJ RADA zaposlenika format['hh']");
            //ob1.endWork = int.Parse(Console.ReadLine());


            //Set name and surname
            //sheet.GetRow(6).Cells[1].SetCellValue(ob1.ime + " " + ob1.prezime);
            ob1.SetNameSurname(sheet);

            //Set starting date
            //sheet.GetRow(8).Cells[1].SetCellValue(ob1.FirstDay());
            ob1.SetFirstDayOfMonth(sheet);

            //Set ending date 
            //sheet.GetRow(8).Cells[4].SetCellValue(ob1.LastDay());
            ob1.SetLastDayOfMonth(sheet);

            //Preset for populating data
            int startingRow = 11;
            DateTime datum = ob1.FirstDay();
            int endingRow = ob1.DaysInMonth() + startingRow;

            for (int i = startingRow; i < endingRow; i++)
            {

                //sheet.GetRow(i).Cells[0].SetCellValue(datum.Date);
                //sheet.GetRow(i).Cells[1].SetCellValue(datum.Date.ToString("ddd"));
                ob1.SetDateAndDay(sheet, i, datum);


                if (datum.DayOfWeek != DayOfWeek.Saturday && datum.DayOfWeek != DayOfWeek.Sunday && !(ob1.holidayCheck(datum))) 
                {
                    //sheet.GetRow(i).Cells[2].SetCellValue(ob1.startWork);
                    ob1.SetStartWork(sheet, i);
                    //sheet.GetRow(i).Cells[3].SetCellValue(ob1.endWork);
                    ob1.SetEndWork(sheet, i);

                    if (ob1.puerperal)
                    {
                        ob1.SetTotalPuerperal(sheet, i);
                    }else
                    {
                        //sheet.GetRow(i).Cells[5].SetCellValue(ob1.TotalWork());
                        ob1.SetTotalWork(sheet, i);
                    }

                }
                datum = datum.AddDays(1);
            }
            string savePath = @"C:\Users\" + Environment.UserName + @"\desktop\\";

            Satnica.saveTemp(workbook, ob1.FileNameCreator(), savePath);

            var excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;

            Microsoft.Office.Interop.Excel.Workbooks books = excel.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook sheet2 = books.Open(savePath + ob1.FileNameCreator());
        }


    }
}

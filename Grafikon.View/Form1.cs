﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Grafikon.Model;
using NPOI.SS.UserModel;

namespace Grafikon.View
{
    public partial class Satnica : Form
    {
        string savePath = @"C:\Users\" + Environment.UserName + @"\desktop\";
        string partIme = string.Empty;
        string partMj = string.Empty;
        string partGodina = string.Empty;

        public Satnica()
        {
            InitializeComponent();
        }

        private void Satnica_Load(object sender, EventArgs e)
        {
            textBoxSavePath.Text = savePath;
            textBoxRVod.Text = "9";
            textBoxRVdo.Text = "17";
            textBoxPoduzece.Text = "GRAFIKON DIZAJN d.o.o";
        }

        public void btn1_Click(object sender, EventArgs e)
        {
            try
            {
                // Create workbook object and pick workbook sheet
                var workbook = Grafikon.Model.Satnica.openTemp();
                ISheet sheet = workbook.GetSheetAt(0);

                // Create satnica object
                Grafikon.Model.Satnica ob1 = new Grafikon.Model.Satnica();

                // Check is ime and prezime field is empty and thorw exception if it is
                if( string.IsNullOrWhiteSpace(textBoxIme.Text) || string.IsNullOrWhiteSpace(textBoxPrezime.Text))
                {
                    throw new ArgumentException("please fill \"Ime\" and \"Prezime\" fields!") ; 
                }
                // Set all data from textbox to Satnica object if is valid data
                ob1.godina = int.Parse(textBoxGodina.Text);
                ob1.mjesec = int.Parse(textBoxMjesec.Text);
                ob1.startWork = int.Parse(textBoxRVod.Text);
                ob1.endWork = int.Parse(textBoxRVdo.Text);
                ob1.ime = textBoxIme.Text.ToUpper();
                ob1.prezime = textBoxPrezime.Text.ToUpper();
                ob1.nazivPoduzeca = textBoxPoduzece.Text.ToUpper();
                ob1.puerperal = radioButtonPorodiljni.Checked;
                ob1.FieldWork = radioButton1đTeren.Checked;
                ob1.vacation = radioButtonGodisnji.Checked;

                // Populate Copnay name
                ob1.SetCompanyName(sheet);

                // Populate Name and Surname field in xls.
                ob1.SetNameSurname(sheet);

                // Populate FirstDay of selected month field in xls.
                ob1.SetFirstDayOfMonth(sheet);

                // Populate LastDay of selected month field in xls.
                ob1.SetLastDayOfMonth(sheet);

                // Seting starting parametr for xls.
                int startingRow = 11;
                int endingRow = ob1.DaysInMonth() + startingRow;
                DateTime datum = ob1.FirstDay();

                // logic for populating data
                for (int i = startingRow; i < endingRow; i++)
                {
                    // Populate Date and day Column in xls.
                    ob1.SetDateAndDay(sheet, i, datum);

                    // Check for week days and holiday
                    if (datum.DayOfWeek != DayOfWeek.Saturday && datum.DayOfWeek != DayOfWeek.Sunday && !(ob1.holidayCheck(datum)))
                    {
                        // Populate StartWork Column in xls.
                        ob1.SetStartWork(sheet, i);
                        // Populate EndWork Column in xls.
                        ob1.SetEndWork(sheet, i);

                        // Check for puerperal
                        if (ob1.puerperal)
                        {
                            // Populate Puerperal Column in xls.
                            ob1.SetTotalPuerperal(sheet, i);
                        }
                        else if (ob1.FieldWork)
                        {
                            // Populate Field Column in xls.
                            ob1.SetFieldWork(sheet, i);
                        }
                        else if (ob1.vacation)
                        {
                            // Populate Vacation in xls.
                            ob1.SetTotalVacation(sheet, i);
                        }
                        else
                        {
                            // Populate TotalWork in xls.
                            ob1.SetTotalWork(sheet, i);
                        }
                    }
                    datum = datum.AddDays(1);
                }

                //Check fileName change
                string FileName;

                if (textBoxFileName.Text != string.Empty)
                {
                    FileName = textBoxFileName.Text +".xls";
                }
                else
                {
                    FileName = ob1.FileNameCreator();
                }

                // Save file to chosen path location and file name
                Grafikon.Model.Satnica.saveTemp(workbook, FileName, savePath);

                // Open file if checkbox is selected
                if (checkBoxXlsOpen.Checked == true)
                {
                    ob1.openFile(savePath + FileName);
                }

                // Output msg box
                MessageBox.Show(" \"" + FileName + "\" je uspijesno kreiran! \n Nalazi se na putanji :\"" + savePath + "\"");
                textBoxIme.Text = string.Empty;
                textBoxPrezime.Text = string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

        }

        private void textBoxIme_Leave(object sender, EventArgs e)
        {
            partIme = textBoxIme.Text.ToUpper();
            textBoxFileName.Text = string.Format("{0}-{1}-{2}",partIme,partMj,partGodina);
        }

        private void textBoxGodina_Leave(object sender, EventArgs e)
        {
            partGodina=textBoxGodina.Text ;
            textBoxFileName.Text = string.Format("{0}-{1}-{2}", partIme, partMj, partGodina);
        }

        private void textBoxMjesec_Leave(object sender, EventArgs e)
        {
            partMj = textBoxMjesec.Text;
            textBoxFileName.Text = string.Format("{0}-{1}-{2}", partIme, partMj, partGodina);
        }
    }
}

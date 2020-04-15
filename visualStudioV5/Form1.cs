using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace smartFridge_v02
{
    public partial class Form1 : Form
    {

        public static class Globals
        {
            public static int rowNumber = 8;
            public static int columnNumber = 1;

        }
        public Form1()
        {
            InitializeComponent();
            //if (!serialPort1.IsOpen)
            //{
            //    tbMessages.Text = "nothing located in port";
            //    serialPort1.Open();
            //    tbMessages.Text = "port opened";
            //}
            //else
            //{
            //    tbMessages.Text = "port is busy rn";
            //}

            // Open up an excel file next
            Object oExcel = new Object();


        }

        private string rxString;
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            rxString = serialPort1.ReadExisting();
            this.Invoke(new EventHandler(displayText));
        }

        private void displayText(object o, EventArgs e)
        {
            tbMessages.Clear();
            tbMessages.AppendText(rxString);
        }

        private void writeText(string str)
        {
            tbMessages.Clear();
            tbMessages.AppendText(str);
        }

        private void buttonAskForItem_Click(object sender, EventArgs e)
        {
            serialPort1.Write("R");
            serialPort1.Write(tbAskForItem.Text);
            tbAskForItem.Clear();
            serialPort1.Write("@");
        }

        private void buttonPutIn_Click(object sender, EventArgs e)
        {
            int foodExists = 0;
            foodExists = checkDatabase(tbPutIn.Text); //returns a "1" if food is in database. Else, zero.

            if (foodExists == 1)
            {
                //do nothing if the food exists already
            }
            else //if the food does not exist, need to write it, and update counter in excel
            {
                //Get current number of filled rows in database
                string databaseCounter;
                databaseCounter = readFromExcel(1, 6, 1);
                int counter = int.Parse(databaseCounter);

                //Write to the next empty row
                writeToExcel(tbPutIn.Text, counter + 1, 1, 1);

                //Update the database counter and write it back
                counter += 1;
                databaseCounter = counter.ToString();
                writeToExcel(databaseCounter, 1, 6, 1); //cell 1,6 is where the counter is stored

            }


            //serialPort1.Write("E");
            //serialPort1.Write(tbPutIn.Text);
            //tbPutIn.Clear();
            //serialPort1.Write("@");
        }

        //checkDirectory searches the excel sheet containing the food database for a match. It returns
        //a "1" if there is a match, otherwise a zero.
        private int checkDatabase(string findThis)
        {
            int foundFood = 0;  //Flag that is set if a matching food name is found

            //Open the excel sheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet4.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);

            //Search the "A" column for the given food, and set the flag
            string checkFood = excelSheet.get_Range("A1", "A1").Value2.ToString(); //THIS PART NEEDS WORK
            if (checkFood == findThis)
            {
                foundFood = 1;
            }
            else
            {
                //leave flag at zero
            }
            excelBook.Close(true);
            excelApp.Quit();

            return foundFood;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            serialPort1.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void viewContents_Click(object sender, EventArgs e)
        {
            serialPort1.Write("V@");
        }

        private void expDates_Click(object sender, EventArgs e)
        {
            serialPort1.Write("D@");
        }


        private void bShowCal_Click(object sender, EventArgs e)
        {
            Form calForm = new Form();
            MonthCalendar mCal = new MonthCalendar();

            calForm.Controls.Add(mCal);

            // Panel panelCal = new Panel();

            // panelCal.Controls.Add(mCal);
            //mCal.Visible = true;
            calForm.ShowDialog();
            tbMessages.Clear();
            tbMessages.AppendText(mCal.SelectionStart.ToString("yyyyMMdd"));    //Shows date in message box
                                                                                //serialPort1.Write("Z");
                                                                                // serialPort1.Write(mCal.SelectionStart.ToString("yyyyMMdd"));
            calForm.Close();

        }

        private void bOpenVB_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet, excelSheet2;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet4.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);
            excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(2);

            //writeText(excelSheet.get_Range("A2", "A2").Value2.ToString());

            string entry = "testing testing";
            excelSheet.Cells[Globals.rowNumber, Globals.columnNumber] = entry;
            excelSheet2.Cells[Globals.rowNumber, Globals.columnNumber] = entry;
            string read = excelSheet.Cells[1, 1].Value.ToString();
            writeText(read);
            //tbMessages.Clear();
            //tbMessages.AppendText(excelSheet.get_Range("A2", "A2").Value2.ToString());

            excelBook.Close(true);
            excelApp.Quit();
        }

        // This function opens the excel sheet, and writes to the given cell
        private void writeToExcel(string textToWrite, int rowNum, int columnNum, int sheetNum)
        {
            //Open excel worksheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet4.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(sheetNum);

            //Write to the appropriate cell
            excelSheet.Cells[rowNum, columnNum] = textToWrite;

            excelBook.Close(true);
            excelApp.Quit();
        }

        // This function opens the excel sheet, and reads from the given cell
        private string readFromExcel(int rowNum, int columnNum, int sheetNum)
        {
            //Open excel workbook and sheets
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet4.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(sheetNum);

            //Read from appropriate cell
            string readText;
            readText = excelSheet.Cells[rowNum, columnNum].Value.ToString();

            excelBook.Close(true);
            excelApp.Quit();

            return readText;
        }

        private void tbAskForItem_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

﻿using System;
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
        public Form1()
        {
            InitializeComponent();
            //All this is optional text that can be displayed to troubleshoot serial connection
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
        }

        private string rxString; 

        //Writes to message box when anything is received via UART
        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            rxString = serialPort1.ReadExisting();
            this.Invoke(new EventHandler(displayText));
        }

        //Write to message textbox
        private void displayText(object o, EventArgs e)
        {
            //tbMessages.Clear();
            tbMessages.AppendText(rxString);
        }

        //Write to message textbox (more generally)
        private void writeText(string str)
        {
            tbMessages.Clear();
            tbMessages.AppendText(str);
        }

        //Referenced when user submits an item to receive from the fridge
        private void buttonAskForItem_Click(object sender, EventArgs e)
        { 
            GrabFood(tbAskForItem.Text);
            tbAskForItem.Clear();
        }

        //Referenced when user submits an item to put into fridge
        private void buttonPutIn_Click(object sender, EventArgs e)
        {
            //Prompt the user to input an expiration date
            string dateEntered;
            dateEntered = showCal();
            tbMessages.AppendText(dateEntered);

            //Enter food in database if not there
            int foodExists;
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

            //place food item in fridge
            PlaceFood(tbPutIn.Text, dateEntered);


            //serialPort1.Write("E");
            //serialPort1.Write(tbPutIn.Text);
            tbPutIn.Clear();
        }

        //checkDatabase searches the excel sheet containing the food database for a match. It returns
        //a "1" if there is a match, otherwise a zero.
        private int checkDatabase(string findThis)
        {
            int foundFood = 0;  //Flag that is set if a matching food name is found

            //Open the excel sheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1);

            //Search the "A" column for the given food, and set the flag
            //Get the database counter from excel
            string databaseCounter;
            databaseCounter = readFromExcel(1, 6, 1);
            int counter = int.Parse(databaseCounter);

            //Search the "A" column for the given food, and set the flag
            string checkFood;
            for (int i = 2; i <= counter; i++)
            {
                checkFood = excelSheet.Cells[i, 1].Value.ToString();
                if (checkFood == findThis)
                {
                    foundFood = 1;
                }
                else
                {
                    //leave flag at zero
                }
            }
            excelBook.Close(true);
            excelApp.Quit();

            return foundFood;
        }

        //Stores food in fridge database
        private void PlaceFood(string foodItem, string date)
        {
            //open excel sheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(2);

            //check for lowest open space
            bool openSpace = false;
            string food, x, y, z;
            int pos = 2;
            while(openSpace == false && pos < 11)
            {
                food = excelSheet.Cells[pos, 1].Value.ToString();
                if(food == "0")
                {
                    openSpace = true;
                    //write food item to open space
                    excelSheet.Cells[pos, 1] = foodItem;
                    excelSheet.Cells[pos, 2] = date;
                    x = excelSheet.Cells[pos, 4].Value.ToString();
                    y = excelSheet.Cells[pos, 5].Value.ToString();
                    z = excelSheet.Cells[pos, 6].Value.ToString();
                    //serialPort1.Write("D"+x+y+z+"@");
                }
                pos++;
            }
            excelBook.Close(true);
            excelApp.Quit();
        }

        //Removes food from fridge database, sends coordinates to microcontroller
        private void GrabFood(string foodItem)
        {
            //open excel sheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(2);

            //check fridge for desired food
            int pos = 2;
            string x, y, z;
            string food;
            bool foundFood = false;
            while(foundFood == false && pos<11)
            {
                food = excelSheet.Cells[pos, 1].Value.ToString();
                if (food == foodItem)
                {
                    foundFood = true;
                    excelSheet.Cells[pos, 1] = "0";
                    excelSheet.Cells[pos, 2] = "e";
                    x = excelSheet.Cells[pos, 4].Value.ToString();
                    y = excelSheet.Cells[pos, 5].Value.ToString();
                    z = excelSheet.Cells[pos, 6].Value.ToString();

                  //  serialPort1.Write("P"+x+y+z+"@");
                }
                pos++;
            }

            //check to see if there is any food above what you want
            if (foundFood)
            {
                int count = 0;
                int movePos = pos-1;
                string col1;
                col1 = excelSheet.Cells[movePos, 4].Value.ToString();
                for (int i = pos; i < 11; i++)
                {
                    string foods, col2;
                    foods = excelSheet.Cells[i, 1].Value.ToString();
                    col2 = excelSheet.Cells[i, 4].Value.ToString();
                    if (foods != "0" && col1 == col2)
                    {
                        count++;
                        excelSheet.Cells[movePos, 1] = foods;
                        excelSheet.Cells[i, 1] = "0";
                        movePos = i;
                    }
                }
                //serialPort1.Write("C");
                //serialPort1.Write(count.ToString());
                //serialPort1.Write("@");
            }
            else
            {
                writeText("It aint there");
            }

        excelBook.Close(true);
        excelApp.Quit();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            serialPort1.Close();
        }

        //Prints contents of fridge to message textbox, when button pressed by user
        private void viewContents_Click(object sender, EventArgs e)
        {
            //Open excel workbook and sheets
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(2);

            //Read from each cell and send to textbox
            tbMessages.Clear();
            string readFood, readDate;
            for(int i=2; i<=10; i++) //"i" should check up to max number of boxes
            {
                readFood = excelSheet.Cells[i, 1].Value.ToString();
                readDate = excelSheet.Cells[i, 2].Value.ToString();
                if (readFood == "0") //cell stores no real food
                {
                    //Do nothing
                }
                else //cell stores a real food
                {
                    tbMessages.AppendText(readFood + "     ");
                    tbMessages.AppendText(readDate + "\r\n");
                }
            }
            excelBook.Close(true);
            excelApp.Quit();
        }

        //Shows pop-up calendar
        private string showCal()
        {
            //Open up a new form containing a calendar
            Form calForm = new Form();
            MonthCalendar mCal = new MonthCalendar();
            System.Windows.Forms.TextBox box = new System.Windows.Forms.TextBox();
            box.AppendText("Choose an expiration date");

            calForm.Controls.Add(mCal);
            calForm.Controls.Add(box);
            box.Location = new System.Drawing.Point(0, 175);
            box.Size = new System.Drawing.Size(200, 100);
            calForm.ShowDialog();

            //Store the selected value
            string dateChosen;
            dateChosen = mCal.SelectionStart.ToString("MM/dd/yyyy");

            //calForm.Hide();
            calForm.Close();
            return dateChosen;
        }

        // This function opens the excel sheet, and writes to the given cell
        private void writeToExcel(string textToWrite, int rowNum, int columnNum, int sheetNum)
        {
            //Open excel worksheet
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook excelBook;
            Worksheet excelSheet;
            string curDir = Directory.GetCurrentDirectory().ToString();
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
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
            excelBook = excelApp.Workbooks.Open(curDir + @"\testsheet1.xlsx");
            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(sheetNum);

            //Read from appropriate cell
            string readText;
            readText = excelSheet.Cells[rowNum, columnNum].Value.ToString();

            excelBook.Close(true);
            excelApp.Quit();

            return readText;
        }

        //Sends a predetermined character to the microcontroller, allowing for motor movement to continue
        private void continueButton_Click(object sender, EventArgs e)
        {
            tbMessages.Clear();
            serialPort1.Write("C@");
        }
    }
}

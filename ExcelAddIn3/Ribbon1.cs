﻿using System;
using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
//using System.Xaml;

using System.Drawing;

using System.Diagnostics;


using ExcelAddIn2.Models;


namespace ExcelAddIn2
{
    public partial class Ribbon1
    {

        //struct has to be global for some reason
        struct OneColumnMap
        {
            public int SAPosition;
            public string SAHead;
            public int CMPosition;
            public string CMHead;
            public bool Required;
            public string defaultValue;
            public bool mapDB;
            public bool SARequired;
            public string Note;
            public string Definition;
            public string CMSource;
        }

        OneColumnMap thisColumnMap;


        public static int nbrFatalErrors = 0;   //made static so ThisAddin can see it

        public const bool EBayImplemented = false;

        //define exterbnal function to get excel app process id as needed to kill zombie processes when using interop
        //see https://stackoverflow.com/questions/8490564/getting-excel-application-process-id
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        //******************************************************************************************************************************************
        //Get reference to current sheet (with add-in) that will hold resulting SA Lot file
        //*****************************************************************************************************************************************


        ArrayList headingsMap = new ArrayList();

        //Made active sheet global
        //Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

       
        //----------------------------------------------------------------------------------------------------------------------------------

        private void btnLoadCatMast_Click(object sender, RibbonControlEventArgs e)
        {

            //DialogResult dialogResultx = MessageBox.Show("Cancel out", "SHeadings Check", MessageBoxButtons.YesNo);
            //if (dialogResultx == DialogResult.No)
            //{
            //    return;
            //}

            string machineName = Environment.MachineName;

            //Console.WriteLine("Into it");
            //System.Diagnostics.Debug.WriteLine("Fuck you");

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS.UsedRange;
            int thisRowCount = thisRange.Rows.Count;
            int thisColCount = thisRange.Columns.Count;

            thisWS.Name = "Kelleher Lang Lots";

            //thisWS.Protect(UserInterfaceOnly: true, AllowFiltering: true, AllowSorting: true);
            thisWS.Unprotect();

            thisRange.Clear();
            thisRange = thisWS.UsedRange;
            thisRowCount = thisRange.Rows.Count;
            thisColCount = thisRange.Columns.Count;

            nbrFatalErrors = 0;


            //** Clear Comments, Values and color because reloading data
            //for (int r = 1; r <= thisRowCount; r++)
            //{
            //    for (int c = 1; c <= thisColCount; c++)
            //    {
            //        thisWS.Cells[r, c].ClearComments();
            //        thisWS.Cells[r, c].Clear();
            //        thisWS.Cells[r, c].Interior.Color = Color.Transparent;
            //    }
            //}


            //1)BUILD ARRAY/INDEX OF EXCEL SA HEADING TO CM HEADING AND POPULATE SA HEADING FROM DATABASE

            //******************************************************************************************************************************************************
            //1.A. BUILD ARRAY (OF STRUCT: HeadingColumnPositions) TO MAP FROM-TO EXCEL COLUMN HEADINGS USING Table ExcelHeadingMap. At same time populate SA lot file 
            //column headings in current spreadsheer. 
            //*****************************************************************************************************************************************************

            //******************************************************************************************************************************************
            //Get (cATALOG mASTER) FILE TO PULL LOT DATA FROM
            //*****************************************************************************************************************************************
            //TODO: INTERCEPT FILE SAVE/CLOSE SO WARN ABOUT SAVE BAD FILE?
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //openFileDialog1.Filter = "Excel Files|*.xls";   //TODO: ADD xlsx filter
            openFileDialog1.Filter = "xlsx Files|*.xlsx|xls Files|*.xls";
            openFileDialog1.Title = "Select Lang file to format for Kelleher";
            openFileDialog1.Multiselect = true;
            string filename;

            //Open file selection dialog - if canceled out just return. Otherwise perform all processing to suck in the selected Catalog Master 
            //lot export file
            if (openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }

            #region Open/Close SQL Connection and load SimpleAuction headings

            /*
             SQL to retrieve Table ExcelHeadingMap ordered by SA Column Position (starting with 1)
                         /*
             dbo.ExcelHeadingMap
                  [SAColumnNbr] [int] NOT NULL,
                  [SAHeading] [varchar] (100) NOT NULL,
                  [CMColumnNbr] int Null,
                  [CMHeading] [varchar] (100) NULL
             */

            //string msg = "";  //Trace msg

            //OneColumnMap thisColumnMap;  --> made global
            
            LoadHeadingMap();  //loads the internal table - NOT the spreadsheet
  

            //TODO: Populate first row current column with SA heading and comment describing how it is derived
            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //if (thisColumnMap.CMHead != null)
            //{
            //    thisWS.Cells[1, thisColumnMap.SAPosition].ClearComments();
            //    thisWS.Cells[1, thisColumnMap.SAPosition].AddComment("Pulled from Catalog Master field: " + thisColumnMap.CMHead + ".  " + thisColumnMap.Note);
            //}

            foreach (OneColumnMap map in headingsMap)
            {
                ((Excel.Range)thisWS.Cells[1, map.SAPosition]).Value = map.SAHead;
                thisWS.Cells[1, map.SAPosition].ClearComments();

                if (map.CMHead != "")
                {
                    Excel.Comment comment = thisWS.Cells[1, map.SAPosition].AddComment("Pulled from Catalog Master field: " + map.CMHead + ".  " + map.Note);
                    comment.Shape.TextFrame.AutoSize = true;
                }
                else
                {
                    Excel.Comment comment = thisWS.Cells[1, map.SAPosition].AddComment(map.Note);
                    comment.Shape.TextFrame.AutoSize = true;
                }
            }


            //TODO: should you keep this in?
            //DialogResult dialogResult = MessageBox.Show("Dow you want to continue past headings?", "SHeadings Check", MessageBoxButtons.YesNo);
            //if (dialogResult == DialogResult.No)
            //{
            //    return;
            //}



            //Declare reuseable (per file) variables here
            var fromXlApp = new Excel.Application();
            Excel.Workbook fromXlWorkbook;
            Excel._Worksheet fromXlWorksheet;
            Excel.Range fromXlRange;
            int rowCount;
            int colCount;

            int filecount = 0;


            foreach (string fn in openFileDialog1.FileNames)
            {


                //TODO: CHECK FILE NAME BEFORE OPENING AGAINST ??? TO MAKE SURE IT'S AN ?UNPROCESSED? ?NEW? ?WELL-NAMED? CATALOG MASTER FILE
                //filename = openFileDialog1.FileName;
                filename = fn;

                //MessageBox.Show("For pause - here is file name about to process: " + filename);

                //excelApp.StatusBar = String.Format("Processing line {0} on {1}.",rows,rowNum);
                Globals.ThisAddIn.Application.StatusBar = String.Format("Loading file {0}: {1}", filecount + 1, openFileDialog1.SafeFileNames[filecount]);

                //%%%%%%%%%%%%%%% START LOOP HERE

                //var fromXlApp = new Excel.Application();
                fromXlApp.Visible = false; //--> Don't need to see the Catalog Master excel file to suck it in
                fromXlWorkbook = fromXlApp.Workbooks.Open(filename); //this is the fully qualified(local) file name
                //fromXlWorkbook = fromXlApp.Workbooks.Open(@"C:\Users\Nicholas\Documents\My Documents\Describing Development\Excel SubProject\Catalog Master Upload Files\Sale 619 using SQL.xlsx");

                 Process fromPid = GetExcelProcess(fromXlApp);

                //int x = fromXlWorkbook.Sheets.Count;
                fromXlWorksheet = (Excel.Worksheet)fromXlWorkbook.Sheets[1];            //TODO: make sure only one worksheet???
                fromXlWorksheet.Activate();
                fromXlRange = fromXlWorksheet.UsedRange;


                //MessageBox.Show("CM (from) file should be open now ... begin data map/load from CM to current SA lot spreadsheet");
                //Globals.ThisAddIn.Application.StatusBar = "processing file";
                Cursor.Current = Cursors.WaitCursor;

                //TODO: MAKE SURE COUNTS ARE NOT ENTIRE WORKSHEET
                rowCount = fromXlRange.Rows.Count;
                colCount = fromXlRange.Columns.Count;

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!
                //1) Walk table ExcelHeadingMap 

                //string msg = "";


                //OLD WAY RELIED ON KNOWING COLUMN NUMBER - NEW WAY WILL MATCH ON NAME ONLY
                ////WALK THE TO SPREADSHEET ROWS AND POPULATE WITH FROM SPREADSHEET VALUES BASED ON MAPPING (ExcelHeadingMap)
                //for (int r = 2; r <= rowCount; r++)         //r = TO ROW TO FILL - WALK HEADING MAP ARRAY TO LEARN COLUMNS TO COPY
                //{
                //    //SEE IF ANY HEADINGS ARE MAPPED FOR THIS "TO" ROW
                //    foreach (OneColumnMap map in headingsMap)
                //    {
                //        if (map.SAHead != "" && map.CMHead == "" && map.defaultValue == "")
                //        {
                //            continue;
                //        }
                //        else
                //            if (map.SAHead != "" && map.CMHead != "")           //if CM heading mapped to SA Heading move the cm spreadsheet value
                //        {
                //            //use the current row in the "TO" spreadsheet- (outer loop)
                //            //NOTE: this will bring mapped fields over as well. MAPPING will occur in ValidateSpreadsheet();
                //            ((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE spreadsheet to spreadsheet VALUE GET'S MOVED!!!

                //        }
                //        else
                //                if (map.SAHead != "" && map.CMHead == "" && map.defaultValue != "")  //otherwise, if there is a default value stuff it into the sa column
                //        {
                //            //NOTE: Default value will trump mapping
                //            ((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = map.defaultValue; //THIS IS WHERE load default value from ExcelHeadingMap!!!
                //        }
                //    }
                //}

                int SAHeadCol = 0;
                int CMHeadCol = 0;

                //colCount is from worksheet column count
                int colsToInspect = thisColCount;
                if (colCount > thisColCount)
                {
                    colsToInspect = colCount;
                }


                //for (int r = 2; r <= rowCount; r++)         //r = TO ROW TO FILL - WALK HEADING MAP ARRAY TO LEARN COLUMNS TO COPY
                //{
                //SEE IF ANY HEADINGS ARE MAPPED FOR THIS "TO" ROW
                foreach (OneColumnMap map in headingsMap)
                {
                    //test
                    //if (map.SAHead == "OriginalSymbols" || map.CMHead == "Stamp Symbols")
                    //{
                    //    MessageBox.Show("symbols!");
                    //}

                    if (map.SAHead != "" && map.CMHead == "" && map.defaultValue == "")  //TODO: what is this???????????
                    {
                        continue;
                    }
                    else
                        if (map.SAHead != "" && map.CMHead != "")           //if CM heading mapped to SA Heading move the cm spreadsheet value
                    {
                        //use the current row in the "TO" spreadsheet- (outer loop)
                        //NOTE: this will bring mapped fields over as well. MAPPING will occur in ValidateSpreadsheet();
                        //NOTE: This assumes column headings are uniquie within spreadsheets
                        //for (int i = 1; i < 256; i++)   //TODO: NEED TO LIMIT TO WHATEVER HIGHER: NBR COLUMNS IN TO OR FROM WORKSHEET
                        //for (int i = 1; i < colsToInspect; i++)   //TODO: NEED TO LIMIT TO WHATEVER HIGHER: NBR COLUMNS IN TO OR FROM WORKSHEET
                        for (int i = 1; i < 256; i++)   //TODO: NEED TO LIMIT TO WHATEVER HIGHER: NBR COLUMNS IN TO OR FROM WORKSHEET
                        {
                            if (thisWS.Cells[1, i].Value == map.SAHead)
                            {
                                SAHeadCol = i;
                            }

                            //int testcolCount = fromXlRange.Columns.Count;
                            //MessageBox.Show("fromXlWorksheet.Cells[1, 10].Text: " + fromXlWorksheet.Cells[1, 10].Text.ToString());
                            //MessageBox.Show("fromXlRange.Cells[1, 10].Text: " + fromXlRange.Cells[1, 10].Text.ToString());
                            //MessageBox.Show("map.CMHead: " + map.CMHead);

                            //Excel.Range testXlRange;
                            //testXlRange = fromXlWorksheet.UsedRange;
                            //if (fromXlRange.Cells[1, i].Value == map.CMHead)  
                            //if (testXlRange.Cells[1, i].Value == map.CMHead)  

                            //var x = fromXlWorksheet.Columns[1].Value;

                            //((Excel.Range)thisWS.Cells[1, 1]).Value = ((Excel.Range)fromXlWorksheet.Cells[2, 1]).Value; --> fromXlWorksheet!!

                            //string xx = ((Excel.Range)fromXlWorksheet.Cells[1, i]).Text.ToString();

                            //if (((Excel.Range)fromXlWorksheet.Cells[1, i]).Value == map.CMHead)
                            if (((Excel.Range)fromXlWorksheet.Cells[1, i]).Value2 == map.CMHead)
                            //if (fromXlRange.Cells[1, i].Value == map.CMHead)  
                            {
                                CMHeadCol = i;
                            }

                            if (SAHeadCol != 0 && CMHeadCol != 0)
                            {
                                break;
                            }
                        }

                        //if (SAHeadCol != 0 && CMHeadCol != 0 && map.CMHead == "Stamp Symbols")
                        //{
                        //    string z = fromXlRange.Cells[2, CMHeadCol].Value2;
                        //    string p = ((Excel.Range)fromXlWorksheet.Cells[2, CMHeadCol]).Value2;
                        //}

                        if (SAHeadCol != 0 && CMHeadCol != 0)       //TODO: LOG IF DB HAS COLUMN MAPPINGS BUT YOU CAN'T FIND COLUMS IN FROM OR TO SPREADSHEET
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = (fromXlRange.Cells[r, CMHeadCol].Value2);  //%%here it is

                                

                            switch (((Excel.Range)thisWS.Cells[r, SAHeadCol]).Text.ToString())
                            {
                                case "FALSE":
                                    ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = "0";
                                    break;
                                case "TRUE":
                                    ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = "1";
                                    break;
                                case "NULL":
                                    ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = "";
                                    break;
                                case "\\N":
                                    ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = "";
                                    break;
                                default:
                                    break;
                            }
                        }
                        }

                        SAHeadCol = 0;
                        CMHeadCol = 0;
                    }
                    else
                    if (map.SAHead != "" && map.CMHead == "" && map.defaultValue != "")  //otherwise, if there is a default value stuff it into the sa column
                    {
                        //NOTE: Default value will trump mapping
                        for (int i = 1; i < 256; i++)
                        {
                            if (thisWS.Cells[1, i].Value2 == map.SAHead)
                            {
                                SAHeadCol = i;
                            }
                        }

                        if (SAHeadCol != 0)
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value2 = map.defaultValue;
                            }
                        }

                        SAHeadCol = 0;
                    }
                    //}
                }



                //TODO: cleanup EXCEL and connections
                //GC.Collect();
                //GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                //Marshal.ReleaseComObject(fromXlRange);
                //Marshal.ReleaseComObject(fromXlWorksheet);  ---> if you do this then workbook.close() fails
                // Marshal.ReleaseComObject(fromXlApp);

                ///****************************************OLD WAY*************************************
                //close and release
                //fromXlWorkbook.Close();

                //quit and release
                //fromXlApp.Quit();
                ///****************************************OLD WAY*************************************



                //*********************************************%%%%%%%%%%%%%****************************************************************************************************************
                //* This will check for missing values and also map database (id) values for category, auction/sale, consignor, consignment
                //*************************************************************************************************************************************************************

                //MessageBox.Show("about to validate");

                //ValidateSpreadsheet(thisWS, true);
                ValidateSpreadsheet(true);

                //((Excel.Range)thisWS.Cells[1, 1]).Value = "fUCKYOU";
                //((Excel.Range)thisWS.Cells[1, 1]).Value = "fUCKYOU1";
                //((Excel.Range)thisWS.Cells[1, 2]).Value = "fUCKYOU2";
                //((Excel.Range)thisWS.Cells[1, 3]).Interior.Color = Color.Azure;
                //((Excel.Range)thisWS.Cells[1, 3]).Interior.Color = Color.Red;
                //((Excel.Range)thisWS.Cells[1, 4]).Value2 = "Fuckyou3";
                //((Excel.Range)thisWS.Cells[1, 5]).Interior.Color = Color.Beige;

                //((Excel.Range)thisWS.Cells[1, 6]).ClearComments();
                //thisWS.Cells[1, 6].AddComment("WTF!!!!!!!!!");
                //((Excel.Range)thisWS.Cells[1, 6]).ClearComments();
                //thisWS.Cells[1, 6].AddComment("WTF2!!!!!!!!!");



                ////thisRange.Cells[1, 2].Value2 = "FUCKYOUtoo";
                ////Marshal.ReleaseComObject(thisRange);
                ////thisRange = null;
                ////thisRange.Cells[1, 2].Interior.Color = Color.Blue;
                ////((Excel.Range)thisWS.Cells[1, 3]).AddComment("fUCKYOU");
                //// ((Excel.Range)thisWS.Cells[1, 3]).Interior.Color = Color.Aqua; 

                ////((Excel.Range)thisWS2.Cells[r, c]).Value = newSymbols;
                ////((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuck");




                //****************************************new way *************************************
                // Get rid of everything - close Excel
                while (Marshal.ReleaseComObject(fromXlRange) > 0) { }
                fromXlRange = null;
                while (Marshal.ReleaseComObject(fromXlWorksheet) > 0) { }
                fromXlWorksheet = null;
                while (Marshal.ReleaseComObject(fromXlWorkbook) > 0) { }
                fromXlWorkbook = null;
                // //while (Marshal.ReleaseComObject(sheets) > 0) { }
                // //sheets = null;


                //// GC();
                fromXlApp.Quit();
                while (Marshal.ReleaseComObject(fromXlApp) > 0) { }
                fromXlApp = null;
                //// GC();


                fromPid.Kill();  //this is needed to get rid of zombie excel processes, so from excel isn't locked for editing after app closes


                //****************************************new way *************************************

                filecount += 1;

            }
            //%%%% end File Name loop



            //thisRowCount = thisRange.Rows.Count;
            //thisColCount = thisRange.Columns.Count;
            ////thisWS.Protect();  --> this protects the whole sheet
            ////https://stackoverflow.com/questions/44883664/how-to-lock-specific-rows-and-columns-using-excel-interop-c-sharp
            //for (int r = 1; r <= thisRowCount; r++)
            //{
            //    //for (int c = 1; c <= thisColCount; c++)
            //    //{
            //        thisWS.Cells[r, 1].Locked = false;

            //    //}
            //}
            //thisWS.Protect(UserInterfaceOnly: true);

            //Globals.ThisAddIn.Application.StatusBar = String.Format("All Catalog Master files are loaded and Validation is complete");
            Globals.ThisAddIn.Application.StatusBar = String.Format("All Catalog Master files are loaded and Validation is complete. The Number of errors is: {0}", nbrFatalErrors.ToString());

            Cursor.Current = Cursors.Default;
        }



        public void LoadHeadingMap()
        {

            headingsMap.Clear();

            //SqlConnection sqlConnection1 = new SqlConnection("Data Source=MANCINI-AWARE\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True");

            string conn = String.Empty;
            if (Environment.MachineName == "MANCINI-AWARE")
            {
                conn = "Data Source=MANCINI-AWARE\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True";
            }
            else
            {
                conn = "Data Source=KELLY-FILE1\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True";

            }

            SqlConnection sqlConnection1 = new SqlConnection(conn);

            SqlCommand cmd1 = new SqlCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.Connection = sqlConnection1;
            SqlDataReader reader1;
            cmd1.CommandText = "SELECT SAColumnNbr,SAHeading,CMColumnNbr,CMHeading, Required, DefaultValue, mapDB, SARequired, Note, Definition, CMSource FROM dbo.ExcelHeadingMapLang order by SAColumnNbr";
            //TODO: Get New column Definition and append to column heads


            sqlConnection1.Open();
            reader1 = cmd1.ExecuteReader();


            if (reader1.HasRows)
            {
                while (reader1.Read())
                {



                    thisColumnMap.SAPosition = reader1.GetInt32(0);
                    thisColumnMap.SAHead = reader1.GetString(1);
                    //thisColumnMap.CMPosition = reader1.GetInt32(2);
                    thisColumnMap.CMPosition = (reader1.IsDBNull(2) ? 0 : reader1.GetInt32(2));
                    //thisColumnMap.CMHead = reader1.GetString(3);
                    thisColumnMap.CMHead = (reader1.IsDBNull(3) ? "" : reader1.GetString(3));
                    thisColumnMap.Required = (reader1.IsDBNull(4) ? false : reader1.GetBoolean(4));
                    //thisColumnMap.Required = reader1.GetBoolean(4);
                    thisColumnMap.defaultValue = (reader1.IsDBNull(5) ? "" : reader1.GetString(5));
                    thisColumnMap.mapDB = (reader1.IsDBNull(6) ? false : reader1.GetBoolean(6));
                    thisColumnMap.SARequired = (reader1.IsDBNull(7) ? false : reader1.GetBoolean(7));
                    thisColumnMap.Note = (reader1.IsDBNull(8) ? "" : reader1.GetString(8));
                    thisColumnMap.Definition = (reader1.IsDBNull(9) ? "" : reader1.GetString(9));
                    thisColumnMap.CMSource = (reader1.IsDBNull(10) ? "" : reader1.GetString(10));





                    //        //.Cells[row, column];
                    //        //Excel.Range rg = thisWS.Cells[1, 1];
                    //        //rg.Cells[4, 1] = "FUCK1";
                    //        //rg.Cells[4, 2] = "FUCK2";
                    //        //rg.Cells[4, 3] = "FUCK3";
                    //        //rg.Cells[4, 4] = "FUCK4";
                    //        //rg.Cells[5, 1] = "YOU1";
                    //        //rg.Cells[5, 2] = "YOU2";
                    //        //rg.Cells[5, 3] = "YOU3";
                    //        //rg.Cells[5, 4] = "YOU4";

                    



                    //TODO: THIS MIGHT BE FASTER WITH RAW INT INSTEAD OF STRUCT (BY VALUE) FIELD
                    //headingsMap[thisColumnMap.SAPosition] = thisColumnMap;
                    headingsMap.Add(thisColumnMap);


                    //        //TODO: CHECK FOR ERRORS, GENERATE MSG BOX, HIGHLIGHT ERROS
                    //        //Worksheets("Sheet1").Range("A1").Value = 3.14159
                    //        //---> no compile ---> thisWS.Range("12","15").Value = "12,15: Worksheet.Range function";
                    //        //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                    //        //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                    //        //((Excel.Range)ws.Cells[r, c]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                    //        //((Excel.Range)ws.Cells[r, c]).Style.Name = "Normal"


                    //        //attribute.FollowsAttributeId = reader.GetInt32(2);      // int FollowsAttributeId 
                    //        //attribute.MultiSelectInd = reader.GetBoolean(6);        // bit MultiSelect_Ind


                    //msg = "FROM STRUCT --> SAPosition: " + thisColumnMap.SAPosition.ToString() + "/SAHead: " + thisColumnMap.SAHead + "/CMPosition: " + thisColumnMap.CMPosition.ToString() + "/CMHead: " + thisColumnMap.CMHead;
                    //Trace.WriteLine(msg + "\t");


                    //Trace.WriteLine(msg + "\t");

                    //struct OneColumnMap
                    //{
                    //    public int SAPosition;
                    //    public string SAHead;
                    //    public int CMPosition;
                    //    public string CMHead;
                    //}

                }
            }


            //TODO: NEED TO EDIT SOURCE? I.E. CERT BODY WITH NO GRADE OR YEAR?


            //TODO: YOU MAY WANT TO LEAVE THIS CONNECTION OPEN IF STORING STUFF IN DB
            reader1.Close();
            cmd1.Dispose();
            sqlConnection1.Close();
            sqlConnection1.Dispose();


            #endregion
        }


        public static void GC()
        {
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        //see https://stackoverflow.com/questions/8490564/getting-excel-application-process-id
        Process GetExcelProcess(Excel.Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }


       // private void ValidateSpreadsheet(Excel.Worksheet thisWS2 = null, bool newData = false) //default in case hit Validate from Ribbon
        private void ValidateSpreadsheet(bool newData = false) //default in case hit Validate from Ribbon
        {
            //((Excel.Range)thisWS2.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

  

            Excel.Worksheet thisWS2 = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS2.UsedRange;
            int rowCount = thisRange.Rows.Count;
            int colCount = thisRange.Columns.Count;


            //((Excel.Range)thisWS2.Cells[1, 1]).Value = "fUCKYOUX";
            //((Excel.Range)thisWS2.Cells[1, 1]).Value = "fUCKYOU1X";
            //((Excel.Range)thisWS2.Cells[1, 2]).Value = "fUCKYOU2X";
            //((Excel.Range)thisWS2.Cells[1, 3]).Interior.Color = Color.Azure;
            //((Excel.Range)thisWS2.Cells[1, 3]).Interior.Color = Color.Red;
            //((Excel.Range)thisWS2.Cells[1, 4]).Value2 = "Fuckyou3X";
            //((Excel.Range)thisWS2.Cells[1, 5]).Interior.Color = Color.Beige;

            //((Excel.Range)thisWS2.Cells[1, 6]).ClearComments();
            //thisWS2.Cells[1, 6].AddComment("WTFX!!!!!!!!!");
            //((Excel.Range)thisWS2.Cells[1, 6]).ClearComments();
            //thisWS2.Cells[1, 6].AddComment("WTF2X!!!!!!!!!");


            //return;


            if (rowCount < 2)
            {
                MessageBox.Show("You haven't loaded data yet - please load data first");
                return;
            }

            thisWS2.Unprotect();
            nbrFatalErrors = 0;
            //TODO: 1) need to clear colors if old spreadsheet

            //** Clear Comments and color because re-validating data when data has not been loaded/reloaded
            if (!newData)
            {
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        //thisWS2.Cells[r, c].ClearComments();  --> Leave old comments - the only ones replaced will be when you re-transform
                        //thisWS2.Cells[r, c].Interior.Color = Color.Transparent; --> old way
                        ((Excel.Range)thisWS2.Cells[1, 5]).Interior.Color = Color.Beige;

                    }
                }
            }

            Cursor.Current = Cursors.WaitCursor;
            Globals.ThisAddIn.Application.StatusBar = String.Format("We are now validating and converting fields in the target AMS spreadsheet. Please be patient.");

            //*****************************************************************************************************************************************************
            //* Load up the database driven mapped fields: categor, sale, consignor, consignment
            //****************************************************************************************************************************************************


            //Set up connection for multiple queries
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=MANCINI-AWARE\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True");

            string conn = String.Empty;
            if (Environment.MachineName == "MANCINI-AWARE")
            {
                conn = "Data Source=MANCINI-AWARE\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True";
            }
            else
            {
                conn = "Data Source=KELLY-FILE1\\SQLEXPRESS ;Initial Catalog=Describing;Integrated Security=True";

            }

            //SqlConnection sqlConnection2 = new SqlConnection(conn);


            //SqlCommand cmd2 = new SqlCommand();
            //cmd2.CommandType = CommandType.Text;
            //cmd2.Connection = sqlConnection2;
            //sqlConnection2.Open();

            //SqlDataReader reader2;



            //TODO: MAKE SURE GETTING LAST ROW AND LAST COLUMN

            //****************************************************************************************************************************************************************************
            //* Validate Required fields (indicated in ExcelHeadingMap
            //****************************************************************************************************************************************************************************

          
            // 1) Build list of SA columns that are required 
            ArrayList reqSaColName = new ArrayList();
            ArrayList mapSaColName = new ArrayList();
            foreach (OneColumnMap map in headingsMap)   //TODO: build headingsMap in initialize function in case Validate hit on existing spreadsheet. Also what if unrelated spreadsheet?
            {
                if (map.Required)           //if CM heading mapped to SA Heading
                {
                    reqSaColName.Add(map.SAHead);
                }
                //TODO - IS this used? (mapped)
                if (map.mapDB)           //if CM heading mapped to SA Heading
                {
                    mapSaColName.Add(map.SAHead);
                }

            }

            //get count of required columns to walk
            //int requiredCount = reqSaColName.Count;
            

            ////*********************************************************************************************************************************************************
            ////* Edit for required fields first
            ////*********************************************************************************************************************************************************
            ////System.Type type;

            //for (int r = 2; r <= rowCount; r++)
            //{
            //    for (int c = 1; c <= colCount; c++)
            //    {
            //        // 3) Make sure all required fields are present
            //        foreach (string reqSAName in reqSaColName)
            //        {
            //            //value = thisWS2.Cells[r, c].Value;
            //            //if ((c == reqSACol) && ((thisWS2.Cells[r, c].Value == null) || (valueint == 0))) {

            //            //TODO --> type = (thisWS2.Cells[r, c].Value).GetType;  //TODO: NEED TO CHECK FOR 0 IN SOME FIELDS I.E. SEQUENCE NBR
            //            if (thisWS2.Cells[1, c].Value == reqSAName && thisWS2.Cells[r, c].Value == null)
            //            {
            //                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;

            //                //TODO: add row,column and heading to comment
            //                String txt = thisWS2.Cells[1, c].Value;
            //                thisWS2.Cells[r, c].ClearComments();
            //                thisWS2.Cells[r, c].AddComment(txt + " is required");
            //                //((Excel.Range)ws.Cells[r, c]).Style.Name = "Normal"

            //                nbrFatalErrors++;
            //            }
            //        }
            //    }
            //}

            //****************************************************************************************************************************************************************************
            // Field Mapping exercises - IS Field title specific - no longer relies on column number
            //****************************************************************************************************************************************************************************

            //int SACategoryId = 0;
            //string GumCode = "";
            //string PackageCode = "";
            //string EbayCategoryId = "";
            //int SAAuctionId = 0;
            //int SAConsignorId = 0;
            //int SAConsignmentId = 0;
            //int SAPriceGuide1Id = 0;
            //int SAPriceGuide2Id = 0;
            //int SALOAProviderId = 0;
            //int SAShippingCategoryId = 0;

            //key column numbers used in xformations
            //int packageUnitsCol = 0;
            //int pubEstInternalCol = 0;
            //int nbrInternetPhotosCol = 0;
            //int nbrCatalogPhotosCol = 0;
            //int formatCol = 0;
            //int catalog1Col = 0;
            //int collectionTypeCol = 0;
            //int originalSymbolCol = 0;
            //int gumStampCol = 0;
            //int packageTypeCol = 0;
            //int reserveTypeCol = 0;
            //int consignorNoCol = 0;
            //int consignmentNoCol = 0;
            //int publicEstLowCol = 0;
            //int publicEstHighCol = 0;
            //int lotDescCol = 0;


            //****************
            int itemtitleCol = 0;
            int scottnbrCol = 0;
            int inquoteCol = 0;
            int endnbrCol = 0;
            int hsamCol = 0;
            //****************


            //List of columns that can be maintained directly
            var openColumns = new List<int>();

            //double packageUnits = 0;
            //double nbrInternetPhotos = 0;
            //double nbrCatalogPhotos = 0;
            //string packageUnitsString = "";
            //string formatString = "";

            int intSaleNo = 0;
            string SaleNo = String.Empty;
            //Console.WriteLine("First column count (colCount): {0}", colCount.ToString());
            //System.Diagnostics.Debug.WriteLine("First column count (colCount): {0}", colCount.ToString());
            //string testTitle = "";

            string heading = String.Empty;
            for (int c = 1; c <= colCount; c++)
            {


                //Console.WriteLine("Column: {0} has title {1}", c.ToString(), thisWS2.Cells[1, c].Value);
                //testTitle = thisWS2.Cells[1, c].Value;
                //System.Diagnostics.Debug.WriteLine("Column: {0}  ",c.ToString());
                //System.Diagnostics.Debug.WriteLine("Title: {0}  ", testTitle);


                //Get column nbr of fields needed later for complext transformations. Also note fields that can be manually maintaing to unlock
                //TODO FATAL ERROR IF N/F
                //TODO: MAKE A SWITCH STMT 

                //heading = thisWS2.Cells[1, c].Value;
                heading = thisWS2.Cells[1, c].Value2;
                //if (thisWS2.Cells[1, c].Value == "Sale_No") {
               if (heading == "Item Title")   //this target name will have been mapped previously
                {
                    //get Est_Cons column to be used later for "Public Estimate 1 (Low)" if internet sale 
                    itemtitleCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Package_units")   //this target name will have been mapped previously
                else if (heading == "ScottNbr")   //this target name will have been mapped previously
                {
                    //get Est_Cons column to be used later for "Public Estimate 1 (Low)" if internet sale 
                    scottnbrCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Nbr Photos Internet Actual")
                else if (heading == "InQuote")
                {
                    inquoteCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Nbr Photos Catalog Actual")
                else if (heading == "EndNbr")
                {
                    endnbrCol = c;
                }
                else if (heading == "TAG")
                {
                    hsamCol = c;
                }
               
                



            }
            bool result = int.TryParse(SaleNo, out intSaleNo);  //TODO: TEST THE Saleno RESULT AND ABORT IF NOT INTEGER

            
            //Size and autowrap all columns
            for (int i = 1; i <= colCount; i++) // this will apply it from col 1 to 10
            {
                thisWS2.Columns[i].ColumnWidth = 25;
            }
            thisWS2.Columns[itemtitleCol].ColumnWidth = 75;
            

            thisWS2.Cells.Style.WrapText = true;   //this so comments will recognize line break

            string t = "";
            //string u = thisWS2.Cells[2, originalSymbolCol].Value2;
            //string y = thisWS2.Cells[2, originalSymbolCol].Value;

            //Check fields that require database mapping by name
            //$$$$$$$$$$
            for (int r = 2; r <= rowCount; r++)
                {
                for (int c = 1; c <= colCount; c++)
                {
                    //t = ((Excel.Range)thisWS2.Cells[1, c]).Value2;
                    //t = thisRange.Cells[1, c].Value2; //TODO: comment this
                    t = thisWS2.Cells[1, c].Value2;

                    //System.Diagnostics.Debug.WriteLine("XTitle: " + t);

                    //WARNING!!! this test for nulls will knock out test if value is null - use a default if testing a field that might be null
                    //if ((thisWS2.Cells[1, c].Value != null) && (thisWS2.Cells[r, c].Value != null))   //mapped columns are required so will be red if not provided
                    if (thisWS2.Cells[1, c].Value != null)    //mapped columns are required so will be red if not provided
                    {

                         if (thisWS2.Cells[1, c].Value2 == "ScottNbr" && thisWS2.Cells[r, itemtitleCol].Value2 != null)
                            {
                                string symbols = thisWS2.Cells[r, itemtitleCol].Value2;
                                var imgs = new List<string>();


                                int i = 0;
                                while ((i = symbols.IndexOf("#", i)) != -1)
                                {
                                    // Print out the substring.
                                    //Console.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);
                                    //Debug.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i, 3), i);

                                    //<img src="http://www.kelleherauctions.com/images/mint.gif" align="top">
                                    string imgname = "";        //TODO: SHOULD DEFINE THESE OUTSIDE OF WHILE?
                                    //int start = i + 45;
                                    //int space = symbols.IndexOf(" ", i);
                                    int space = symbols.IndexOfAny(new char[] { ' ', ',', '#' }, i+1);
                                //backslash += 1;
                                //int period = symbols.IndexOf(".", start);
                                //int displace = 0;
                                int length = space - i;

                                if (space != -1 && length > 1)
                                    {
                                        ///displace = period - backslash;

                                        imgname = symbols.Substring(i, length);
                                        //Debug.WriteLine("row: {0} image name {1}", r, imgname);

                                        imgs.Add(imgname);
                                    }

                                    // Increment the index.
                                    i++;
                                }

                            if (imgs.Count > 0)
                            {
                                foreach (string s in imgs)
                                {
                                    thisWS2.Cells[r, c].Value = thisWS2.Cells[r, c].Value + s + ";";
                                }
                            }

                            //**********************************************************************************
                            //String term = string.Empty;
                            //bool inbracket = false;

                            //foreach (char ch in symbols)
                            //{
                            //    if (ch == '<')
                            //    {
                            //        inbracket = true;
                            //        if (term != "")
                            //        {
                            //            imgs.Add(term);
                            //            term = "";
                            //        }

                            //    }
                            //    else
                            //        if (ch == '>')
                            //    {
                            //        inbracket = false;
                            //    }
                            //    else
                            //        if (!inbracket && ch != '/' && ch != '\'' && ch != '\"')
                            //    {
                            //        term += ch;
                            //    }
                            //}

                            //if (term != "")
                            //{
                            //    imgs.Add(term);
                            //    term = "";
                            //}

                            //TODO: SHOULD DEFINE THESE OUTSIDE OF WHILE?
                            //string imgid = String.Empty;
                            //string origimg = String.Empty;

                            //string newSymbols = String.Empty;
                            //string newComment = String.Empty;
                            //bool wasError = false;

                            //if (imgs.Count > 0)
                            //{

                            //origimg = thisWS2.Cells[r, originalSymbolCol].Value2;
                            //((Excel.Range)thisWS2.Cells[r, c]).Value = "";  //clear out value
                            //((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                            //newComment = "Original symbol before xform: " + origimg;

                            //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;



                            //%%%%BEGIN REAL SECTION
                            //foreach (string s in imgs)
                            //{

                            //    /*
                            //    Number sign &#36; 
                            //    Dollar sign &#37; 
                            //    Percent sign &#38; 
                            //    Ampersand &#39; 
                            //    Apostrophe &#40; 
                            //    Left parenthesis &#41; *******
                            //    Right parenthesis &#42; 
                            //    Asterisk &#43;
                            //    */

                            //    if (s == "&#41" || s == "&#41;" || s == "&#40" || s == "&#40;")
                            //    {
                            //        continue;
                            //    }

                            //    string s2 = s.Replace("&#41", "");
                            //    s2 = s2.Replace("&#41;", "");
                            //    s2 = s2.Replace("&#40", "");
                            //    s2 = s2.Replace("&#40;", "");

                            //    cmd2.CommandText = "select amsid from symbol_reference where cmid = '" + s2 + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            imgid = reader2.GetString(0);              //Note EBay id is string until further known
                            //            newSymbols = newSymbols + imgid + ",";
                            //        newComment = newComment + " Mapped to " + imgid + ";";

                            //        //thisWS2.Cells[r, c].ClearComments();
                            //        //thisWS2.Cells[r, c].Interior.Color = Color.BurlyWood;
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.BurlyWood;
                            //        //((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuckfuck");
                            //        //thisWS2.Cells[r, c].AddComment("fuckfuck");

                            //        //thisWS2.Cells[r, c].Value = thisWS2.Cells[r, c].Value + imgid + ";";
                            //        //thisWS2.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                            //        //break;
                            //    }
                            //}
                            //    else
                            //    {
                            //    //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //    //thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //    //TODO: add row,column and heading to comment
                            //    //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //    //thisWS2.Cells[r, c].ClearComments();
                            //    wasError = true;
                            //    newComment = newComment + "Tried to find image for CMid " + s2 + " in table Symbol_Reference. CMid not found - please add to table";
                            //    //thisWS2.Cells[r, c].AddComment("Tried to find image for CMid " + s2 + " in table Symbol_Reference. CMid not found - please add to table");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}

                            //if (newSymbols.Length > 0)
                            //{
                            //    newSymbols = newSymbols.TrimEnd(',');
                            //    //thisWS2.Cells[r, c].Value = newSymbols;
                            //    ((Excel.Range)thisWS2.Cells[r, c]).Value2 = newSymbols;
                            //    //thisRange.Cells[r, c].Value2 = "bbb";
                            //    //thisRange.Cells[r, c].Interior.Color = Color.Blue;

                            //    //((Excel.Range)thisWS2.Cells[r, c]).Value = newSymbols;
                            //    //((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuck");
                            //}

                            ////thisWS2.Cells[r, c].AddComment(newComment); //--> OLD WAY STOPPED WORKING
                            ////((Excel.Range)fromXlWorksheet.Cells[1, i]).Value --> PROTOTYPE
                            //thisWS2.Cells[r, c].ClearComments();
                            //Excel.Comment comment = ((Excel.Range)thisWS2.Cells[r, c]).AddComment(newComment);
                            //comment.Shape.TextFrame.AutoSize = true;
                            //if (wasError)
                            //{
                            //    thisWS2.Cells[r, c].Interior.Color = Color.Red;
                            //    nbrFatalErrors++;
                            //}
                            //else
                            //{
                            //    thisRange.Cells[r, c].Interior.Color = Color.Blue;

                            //}

                            //%%%%%END REAL SECTION
                        }
                        else if (thisWS2.Cells[1, c].Value2 == "InQuote" && thisWS2.Cells[r, itemtitleCol].Value2 != null)
                        {
                            string symbols = thisWS2.Cells[r, itemtitleCol].Value2;
                            //var imgs = new List<string>();


                            int startquote = 0;
                            int endquote = 0;

                            startquote = symbols.IndexOf('"', 0);
                            if (startquote != -1)
                            {
                                endquote = symbols.IndexOf('"', startquote + 1);

                                if (endquote != -1)
                                {
                                    string quoted = "";        //TODO: SHOULD DEFINE THESE OUTSIDE OF WHILE?
                                    int length = endquote - startquote;
                                    quoted = symbols.Substring(startquote + 1, length -1);
                                    thisWS2.Cells[r, c].Value = quoted;
                                }
                            }
                        }
                        else if (thisWS2.Cells[1, c].Value2 == "EndNbr" && thisWS2.Cells[r, itemtitleCol].Value2 != null)
                        {
                            string symbols = thisWS2.Cells[r, itemtitleCol].Value2;


                            int length = symbols.Length;
                            int lastspace = symbols.LastIndexOf(" ");
                            int wordlength = 0;

                            if (lastspace != -1 )
                            {
                                wordlength = length - (lastspace + 1);
                                //if (wordlength == 4 && symbols.Substring(lastspace + 1, 4) == "HSAM")
                                if (wordlength < 5 || wordlength > 7)
                                {
                                    string tagvalue = symbols.Substring(lastspace + 1, wordlength);

                                    int secondlastspace = symbols.LastIndexOf(" ", lastspace -1);
                                    if (secondlastspace != -1)
                                    {
                                        wordlength = lastspace - (secondlastspace + 1);
                                        if (wordlength < 8 && wordlength > 4)
                                        {
                                            string itemid = symbols.Substring(secondlastspace + 1, wordlength);
                                            thisWS2.Cells[r, c].Value = itemid;

                                            thisWS2.Cells[r, hsamCol].Value = tagvalue;


                                        }
                                    }
                                }
                                else
                                {
                                    if (wordlength < 8 && wordlength > 4)  //VS, WLB, BKEY, HSAM, DG, awol
                                    {
                                        string itemid = symbols.Substring(lastspace + 1, wordlength);
                                        thisWS2.Cells[r, c].Value = itemid;
                                    }
                                }
                            }



                            
                        }
                    }
                    }
                    
            }


           


            //*********************************************************************************************************************************************************
            //* Edit for required fields after xforms done now because you're not pre-stuffing xform fields so can re-validate cleanly
            //*********************************************************************************************************************************************************
            //System.Type type;

            for (int r = 2; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    // 3) Make sure all required fields are present
                    foreach (string reqSAName in reqSaColName)
                    {
                        //value = thisWS2.Cells[r, c].Value;
                        //if ((c == reqSACol) && ((thisWS2.Cells[r, c].Value == null) || (valueint == 0))) {

                        //TODO --> type = (thisWS2.Cells[r, c].Value).GetType;  //TODO: NEED TO CHECK FOR 0 IN SOME FIELDS I.E. SEQUENCE NBR
                        if (thisWS2.Cells[1, c].Value2 == reqSAName && thisWS2.Cells[r, c].Value2 == null)
                        {
                            ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;

                            //TODO: add row,column and heading to comment
                            String txt = thisWS2.Cells[1, c].Value2;
                            ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                            thisWS2.Cells[r, c].AddComment(txt + " is required");

                            nbrFatalErrors++;
                        }
                    }
                }
            }




            // In Excel, you can only effectively lock cells if you lock the worksheet.What you do is:
            // Mark the cell ranges you don't want to lock as Locked = False
            // Then protect the sheet using sheet.Protect(UserInterfaceOnly: true).
            // thisWS2.Protect();  --> this protects the whole sheet
            //https://stackoverflow.com/questions/44883664/how-to-lock-specific-rows-and-columns-using-excel-interop-c-sharp
                foreach (int c in openColumns)
            {
                for (int r = 2; r <= rowCount; r++)
                {
                    //thisWS2.Cells[r, c].Locked = true;
                    //thisWS2.Cells[r, c].Locked = false;
                    ((Excel.Range)thisWS2.Cells[r, c]).Locked = false;
                }
            }
            //Excel.Range newRange = thisWS2.UsedRange;

            thisWS2.Activate();
            thisWS2.Application.ActiveWindow.SplitRow = 1;
            thisWS2.Application.ActiveWindow.FreezePanes = true;

            // Now apply autofilter: true allows the user to set filters on the protected worksheet. 
            // Users can change filter criteria but can not enable or disable an autofilter. 
            // Users can set filters on an existing autofilter. 
            Excel.Range firstRow = (Excel.Range)thisWS2.Rows[1];
            firstRow.AutoFilter(1,
                                Type.Missing,
                                Excel.XlAutoFilterOperator.xlAnd,
                                Type.Missing,
                                true);

            thisWS2.Protect(UserInterfaceOnly: true, AllowFiltering: true, AllowSorting: true);



            //AllowFormattingCells: true
            //Contents: false


            

            Globals.ThisAddIn.Application.StatusBar = String.Format("Validation is complete. The Number of errors is: {0}", nbrFatalErrors.ToString());

            Cursor.Current = Cursors.Default;

        }

        //private int GetInvColumn(r)
        //{
        //    for (i=1; i<256; int++)
        //    {
        //        if (thisWS.Cells[1, i].Value = "SerialNumber")
        //        {
        //            return (i);
        //        }
        //    }

        //    return 0;
        //}

      


        private void btnVerify_Click_1(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("You hit the verify button");
            //thisColumnMap = record
            //headingsMap = table of thisColumnMap records

            if (headingsMap.Count == 0)
            {
                LoadHeadingMap();
            }

            //private void ValidateSpreadsheet(Excel.Worksheet thisWS2 = null, bool newData = false)


            ValidateSpreadsheet();

           

        }

        //private void btnSelectCategory_Click_1(object sender, RibbonControlEventArgs e)
        //{
        //    Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;


        //    if (!thisWS.AutoFilterMode)
        //    {
        //        MessageBox.Show("Autofilter must be on to assign category codes ");
        //        return;
        //    }

        //    //MessageBox.Show("da fuck!");
        //    var form1 = new Form1();
        //    form1.Show();

        //}

        //capture the woekbook save event and warn if errors
        //private void WorkbookBeforeSave()
        //{
        //    this.BeforeSave +=
        //        new Excel.WorkbookEvents_BeforeSaveEventHandler(
        //        ThisWorkbook_BeforeSave);
        //}

        //void ThisWorkbook_BeforeSave(bool SaveAsUI, ref bool Cancel)
        //{
        //    if (DialogResult.No == MessageBox.Show("Are you sure you want to " +
        //        "save the workbook?", "Example", MessageBoxButtons.YesNo))
        //    {
        //        Cancel = true;
        //        MessageBox.Show("Save is canceled.");
        //    }
        //}

    }
}

using System;
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

            thisWS.Name = "Kelleher Catalog Lots";

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
            openFileDialog1.Title = "Select Catalog Master Lot Files to format for Simple Auction";
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
                    thisWS.Cells[1, map.SAPosition].AddComment("Pulled from Catalog Master field: " + map.CMHead + ".  " + map.Note);
                }
                else
                {
                    thisWS.Cells[1, map.SAPosition].AddComment(map.Note);
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

                        if (SAHeadCol != 0 && CMHeadCol != 0 && map.CMHead == "Stamp Symbols")
                        {
                            string z = fromXlRange.Cells[2, CMHeadCol].Value2;
                            string p = ((Excel.Range)fromXlWorksheet.Cells[2, CMHeadCol]).Value2;
                        }

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
            cmd1.CommandText = "SELECT SAColumnNbr,SAHeading,CMColumnNbr,CMHeading, Required, DefaultValue, mapDB, SARequired, Note, Definition, CMSource FROM dbo.ExcelHeadingMap order by SAColumnNbr";
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

            SqlConnection sqlConnection2 = new SqlConnection(conn);


            SqlCommand cmd2 = new SqlCommand();
            cmd2.CommandType = CommandType.Text;
            cmd2.Connection = sqlConnection2;
            sqlConnection2.Open();

            SqlDataReader reader2;



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
            string GumCode = "";
            string PackageCode = "";
            //string EbayCategoryId = "";
            //int SAAuctionId = 0;
            //int SAConsignorId = 0;
            //int SAConsignmentId = 0;
            //int SAPriceGuide1Id = 0;
            //int SAPriceGuide2Id = 0;
            //int SALOAProviderId = 0;
            //int SAShippingCategoryId = 0;

            //key column numbers used in xformations
            int packageUnitsCol = 0;
            int pubEstInternalCol = 0;
            int nbrInternetPhotosCol = 0;
            int nbrCatalogPhotosCol = 0;
            int formatCol = 0;
            int catalog1Col = 0;
            int collectionTypeCol = 0;
            //int alphaTextCol = 0;
            int originalSymbolCol = 0;
            int gumStampCol = 0;
            int packageTypeCol = 0;
            int reserveTypeCol = 0;
            int consignorNoCol = 0;
            int consignmentNoCol = 0;
            int publicEstLowCol = 0;
            int publicEstHighCol = 0;

            //List of columns that can be maintained directly
            var openColumns = new List<int>();

            double packageUnits = 0;
            double nbrInternetPhotos = 0;
            double nbrCatalogPhotos = 0;
            string packageUnitsString = "";
            string formatString = "";

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
                if (heading == "Sale_No") {
                    //Get the saleno by finding the Sale_No column and getting first row value
                    SaleNo = thisWS2.Cells[2, c].Text.ToString();
                    continue;
                }   //TODO: Determine if still need to allow SrtOrder to be editable
                //else if (thisWS2.Cells[1, c].Value == "SrtOrder") {
                else if (heading == "SrtOrder") {
                    //Create list of columns that can be directly changed
                    openColumns.Add(c);
                }
                //else if (thisWS2.Cells[1, c].Value == "Est_Cons")   //this target name will have been mapped previously
                else if (heading == "Est_Cons")   //this target name will have been mapped previously
                {
                    //get Est_Cons column to be used later for "Public Estimate 1 (Low)" if internet sale 
                    pubEstInternalCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Package_units")   //this target name will have been mapped previously
                else if (heading == "Package_units")   //this target name will have been mapped previously
                {
                    //get Est_Cons column to be used later for "Public Estimate 1 (Low)" if internet sale 
                    packageUnitsCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Nbr Photos Internet Actual")
                else if (heading == "Nbr Photos Internet Actual")
                {
                    nbrInternetPhotosCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Nbr Photos Catalog Actual")
                else if (heading == "Nbr Photos Catalog Actual")
                {
                    nbrCatalogPhotosCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Format")
                else if (heading == "Format")
                {
                    formatCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Catalog 1 Number")
                else if (heading == "Catalog 1 Number")
                {
                    catalog1Col = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Collection Type")
                else if (heading == "Collection Type")
                {
                    collectionTypeCol = c;
                }
                //else if (heading == "Alpha Text")
                //{
                //    alphaTextCol = c;
                //}
                else if (heading == "OriginalSymbols")
                {
                    originalSymbolCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Gum (Stamps)")
                else if (heading == "Gum (Stamps)")
                {
                    gumStampCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Package_type")
                else if (heading == "Package_type")
                {
                    packageTypeCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Reserve Type")
                else if (heading == "Reserve Type")
                {
                    reserveTypeCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Consignor No.")
                else if (heading == "Consignor No.")
                {
                    consignorNoCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Consignment Number -Incoming")
                else if (heading == "Consignment Number -Incoming")
                {
                    consignmentNoCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Public Estimate 1 (Low)")
                else if (heading == "Public Estimate 1 (Low)")
                {
                    publicEstLowCol = c;
                }
                //else if (thisWS2.Cells[1, c].Value == "Public Estimate 2 (High)")
                else if (heading == "Public Estimate 2 (High)")
                {
                    publicEstHighCol = c;
                }



            }
            bool result = int.TryParse(SaleNo, out intSaleNo);  //TODO: TEST THE Saleno RESULT AND ABORT IF NOT INTEGER


            string t = "";
            string u = thisWS2.Cells[2, originalSymbolCol].Value2;
            string y = thisWS2.Cells[2, originalSymbolCol].Value;

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

                        //if (thisWS2.Cells[1, c].Value == "Consignor")
                        //{
                        //    SAConsignorId = 0;

                        //    cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS2.Cells[r, consignorNoCol].Value + "'";    //mapping step stuffed CM value, so now re-map
                        //    reader2 = cmd2.ExecuteReader();


                        //    if (reader2.HasRows)
                        //    {
                        //        while (reader2.Read())
                        //        {

                        //            //CMCategoryTxt = reader2.GetString(1);
                        //            //SACategoryTxt = reader2.GetString(3);
                        //            //EBCategoryTxt = reader2.GetString(5);

                        //            SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                        //            thisWS2.Cells[r, c].Value = SAConsignorId;
                        //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                        //            thisWS2.Cells[r, c].ClearComments();
                        //            thisWS2.Cells[r, c].AddComment("Mapped CM consignor id: " + thisWS2.Cells[r, consignorNoCol].Value + " to AMS consignor id");

                        //            //break;
                        //        }
                        //    }
                        //    else
                        //    {
                        //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                        //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                        //        //TODO: add row,column and heading to comment
                        //        thisWS2.Cells[r, c].ClearComments();
                        //        thisWS2.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS2.Cells[r, consignorNoCol].Value + " to AMS consignor id - consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");

                        //        //thisWS2.Cells[r, c].Comment[1].Style.Size.AutomaticSize = true;

                        //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;  --> this does not work
                        //        //var comment = thisWS2.Cell("A3").Comment;
                        //        //comment.AddText("This is a very very very very very long line comment.");
                        //        //comment.Style.Size.AutomaticSize = true;

                        //        nbrFatalErrors++;
                        //    }

                        //    reader2.Close();
                        //}
                        //$$$$$$$$$$$
                        //else if (thisWS2.Cells[1, c].Value == "prop_num")
                        //{
                        //    SAConsignmentId = 0;


                        //    cmd2.CommandText = "SELECT SAId FROM dbo.Consignment where CMId = '" + thisWS2.Cells[r, consignmentNoCol].Value + "'";    //mapping step stuffed CM value, so now re-map
                        //    reader2 = cmd2.ExecuteReader();


                        //    if (reader2.HasRows)
                        //    {
                        //        while (reader2.Read())
                        //        {

                        //            //CMCategoryTxt = reader2.GetString(1);
                        //            //SACategoryTxt = reader2.GetString(3);
                        //            //EBCategoryTxt = reader2.GetString(5);

                        //            SAConsignmentId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                        //            thisWS2.Cells[r, c].Value = SAConsignmentId;
                        //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                        //            thisWS2.Cells[r, c].ClearComments();
                        //            thisWS2.Cells[r, c].AddComment("Mapped CM consignment id: " + thisWS2.Cells[r, consignmentNoCol].Value + " to AMS consignment id");

                        //            //break;
                        //        }
                        //    }
                        //    else
                        //    {
                        //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                        //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                        //        //TODO: add row,column and heading to comment
                        //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                        //        thisWS2.Cells[r, c].ClearComments();
                        //        thisWS2.Cells[r, c].AddComment("Tried to map CM consignment id: " + thisWS2.Cells[r, consignmentNoCol].Value + " to AMS consignment id - CM consignment id not found in Consignment table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                        //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                        //        //thisWS2.Cells[r, c].Comment[1].Style.Size.AutomaticSize = true; //TODO: autosize not working?


                        //        nbrFatalErrors++;
                        //    }

                        //    reader2.Close();
                        //}
                        //**************
                        if (thisWS2.Cells[1, c].Value2 == "Est_Low")
                        {
                        if (intSaleNo > 3999)  //Internet sales use internale estimate
                            {
                                thisWS2.Cells[r, c].Value2 = thisWS2.Cells[r, pubEstInternalCol].Value;

                                if (thisWS2.Cells[r, c].Value2 != null)
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("Low Estimate derived from CM Estimate (Internal) because this is an Internet Sale");

                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("Low Estimate derived from CM Estimate (Internal) because this is an Internet Sale");

                                }
                                else
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("Low Estimate could not be derived from CM Estimate (Internal) for Internet sale because field is empty");

                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("Low Estimate could not be derived from CM Estimate (Internal) for Internet sale because field is empty");


                                }
                            }
                            else
                            {
                                thisWS2.Cells[r, c].Value2 = thisWS2.Cells[r, publicEstLowCol].Value2;

                                if (thisWS2.Cells[r, c].Value2 != null)
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("Low Estimate derived from Public Estimate (Low) because this is a Public Sale");

                                    
                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("Low Estimate derived from Public Estimate (Low) because this is a Public Sale");
                                }
                                else
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("Low Estimate could not be derived from CM Public Estimate (Low) for public sale because field is empty");

                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("Low Estimate could not be derived from CM Public Estimate (Low) for public sale because field is empty");

                                }
                            }
                            
                        }
                        //$$$$$
                        else if (thisWS2.Cells[1, c].Value2 == "Est_Real")
                        {
                            //int publicEstLowCol = 0;
                            //int publicEstHighCol = 0;

                            if (intSaleNo > 3999)
                            {
                                thisWS2.Cells[r, c].Value2 = null;  //Estimate high is empty for Internet sales
                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("High Estimate is not used because this is an Internet Sale");

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("High Estimate is not used because this is an Internet Sale");
                            }
                            else
                            {
                                thisWS2.Cells[r, c].Value2 = thisWS2.Cells[r, publicEstHighCol].Value2;

                                if (thisWS2.Cells[r, c].Value != null)
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("High Estimate derived from Public Estimate (High) because this is a Public Sale");

                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("High Estimate derived from Public Estimate (High) because this is a Public Sale");
                                }
                                else
                                {
                                    //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                    //thisWS2.Cells[r, c].ClearComments();
                                    //thisWS2.Cells[r, c].AddComment("High Estimate could not be derived from CM Public Estimate (High) for public sale because field is empty");

                                    ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                    ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                    thisWS2.Cells[r, c].AddComment("High Estimate could not be derived from CM Public Estimate (High) for public sale because field is empty");
                                }
                            }
                        }
                        //$$$$$$$$$$$$$
                        else if (thisWS2.Cells[1, c].Value == "Currency")
                        {

                            if (intSaleNo < 101)  //only Public Hong Kong sales are HK$
                            {
                                thisWS2.Cells[r, c].Value2 = "HK$";

                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Public Hong Kong sale uses HK$");

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Public Hong Kong sale uses HK$");
                            }
                            else
                            {
                                thisWS2.Cells[r, c].Value = "USD$";

                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Public Kelleher, Private Treaty and Internet sales all use USD$");

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Public Kelleher, Private Treaty and Internet sales all use USD$");
                            }

                        }
                        //$$$$$$$$$$$$$$
                        else if (thisWS2.Cells[1, c].Value2 == "Gum" && thisWS2.Cells[r, gumStampCol].Value2 != null) //only go after where CM hase set the gum code
                        {
                            GumCode = "";


                            cmd2.CommandText = "SELECT AMS_Gum_Code FROM dbo.Gum_Codes where CM_Gum_Code = '" + thisWS2.Cells[r, gumStampCol].Value + "'";
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    GumCode = reader2.GetString(0);         //assume it's not red alread because these was a value to lookup
                                    if (GumCode != "")
                                    {
                                        thisWS2.Cells[r, c].Value2 = GumCode;
                                        //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                        ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                        //((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                        //thisWS2.Cells[r, c].AddComment("Does this need a comment?");
                                    }
                                    else
                                    {
                                        //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                        //thisWS2.Cells[r, c].ClearComments();
                                        //thisWS2.Cells[r, c].AddComment("Tried to map CM Gum (Stamp): " + thisWS2.Cells[r, gumStampCol].Value2 + " to AMS gum code CM code found but AMS code blank in Gum_Codes table. Mapping is required - please add mapping to table and re-validate this spreadsheet");

                                        ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                        ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                        thisWS2.Cells[r, c].AddComment("Tried to map CM Gum (Stamp): " + thisWS2.Cells[r, gumStampCol].Value2 + " to AMS gum code CM code found but AMS code blank in Gum_Codes table. Mapping is required - please add mapping to table and re-validate this spreadsheet");


                                        nbrFatalErrors++;
                                    }
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Tried to map CM Gum (Stamp): " + thisWS2.Cells[r, c].Value + " to AMS gum code using table Gum_Codes. Mapping is required - please add mapping to table Gum_Codes and re-validate this spreadsheet");

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Tried to map CM Gum (Stamp): " + thisWS2.Cells[r, c].Value + " to AMS gum code using table Gum_Codes. Mapping is required - please add mapping to table Gum_Codes and re-validate this spreadsheet");


                                nbrFatalErrors++;
                            }

                            reader2.Close();
                        }
                        //** Note it is assumes that all package units will be at least 1
                        else if (thisWS2.Cells[1, c].Value2 == "AMS_Package_Type" && thisWS2.Cells[r, packageTypeCol].Value2 != null) //only go after where CM hase set the gum code
                        {
                            PackageCode = "";
                            packageUnits = 0;

                            cmd2.CommandText = "SELECT AMS_Package_Code FROM dbo.Package_Codes where CM_Package_Code = '" + thisWS2.Cells[r, packageTypeCol].Value2 + "'";
                            reader2 = cmd2.ExecuteReader();

                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    PackageCode = reader2.GetString(0);         //assume it's not red alread because these was a value to lookup
                                    if (PackageCode != "")
                                    {
                                        packageUnits = thisWS2.Cells[r, packageUnitsCol].Value;
                                        if (packageUnits > 1)
                                        {
                                            packageUnitsString = "(" + packageUnits.ToString() + ")";
                                        }
                                        thisWS2.Cells[r, c].Value2 = PackageCode + packageUnitsString;
                                        //thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                                        //thisWS2.Cells[r, c].ClearComments();
                                        //thisWS2.Cells[r, c].AddComment("Mapped CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value + " to AMS Package_Type.");

                                        ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                        ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                        thisWS2.Cells[r, c].AddComment("Mapped CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value2 + " to AMS Package_Type.");


                                    }
                                    else
                                    {
                                        //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                        //thisWS2.Cells[r, c].ClearComments();
                                        //thisWS2.Cells[r, c].AddComment("Tried to map CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value + " to AMS Package_Type. CM code found but AMS code blank in table Package_Codes. Mapping is required - please add mapping to table and re-validate this spreadsheet");

                                        ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                        ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                        thisWS2.Cells[r, c].AddComment("Tried to map CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value2 + " to AMS Package_Type. CM code found but AMS code blank in table Package_Codes. Mapping is required - please add mapping to table and re-validate this spreadsheet");

                                        nbrFatalErrors++;
                                    }
                                    //break;
                                }
                            
                        }
                        else
                        {
                            //thisWS2.Cells[r, c].Interior.Color = Color.Red;
                            //thisWS2.Cells[r, c].ClearComments();
                            //thisWS2.Cells[r, c].AddComment("Tried to map CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value + " to AMS Package Tpe using table Package_Codes. Mapping is required - please add mapping to table Package_Codes and re-validate this spreadsheet");

                            ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                            thisWS2.Cells[r, c].AddComment("Tried to map CM Package_Type: " + thisWS2.Cells[r, packageTypeCol].Value2 + " to AMS Package_Type. CM code found but AMS code blank in table Package_Codes. Mapping is required - please add mapping to table and re-validate this spreadsheet");

                            nbrFatalErrors++;
                        }

                        reader2.Close();
                    }
                        else if (thisWS2.Cells[1, c].Value2 == "AMS_Net_Reserve_Ind")  //this will have been loaded with the reserve type from CM with default to false
                        {
                            if (thisWS2.Cells[r, reserveTypeCol].Text.ToString() == "2" || thisWS2.Cells[r, reserveTypeCol].Text.ToString() == "5")
                            {
                                thisWS2.Cells[r, c].Value2 = "Y";
                            }
                            else
                            {
                                thisWS2.Cells[r, c].Value2 = "N";
                            }

                            //thisWS2.Cells[r, c].ClearComments();
                            //thisWS2.Cells[r, c].AddComment("Mapped CM Reserve type: " + thisWS2.Cells[r, reserveTypeCol].Value + " to AMS Net Reserve Indicator");
                            //thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                            ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                            ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                            thisWS2.Cells[r, c].AddComment("Mapped CM Reserve type: " + thisWS2.Cells[r, reserveTypeCol].Value2 + " to AMS Net Reserve Indicator");


                        }
                        //*********************
                        else if (thisWS2.Cells[1, c].Value == "Derived Example")
                        {
                            //Note this field is defaulted to N
                            nbrInternetPhotos = 0;
                            nbrCatalogPhotos = 0;
                            formatString = "";

                            if (thisWS2.Cells[r, nbrInternetPhotosCol].Value != null)
                            {
                                nbrInternetPhotos = thisWS2.Cells[r, nbrInternetPhotosCol].Value;
                            }
                            if (thisWS2.Cells[r, nbrCatalogPhotosCol].Value != null)
                            {
                                nbrCatalogPhotos = thisWS2.Cells[r, nbrCatalogPhotosCol].Value;
                            }
                            if (thisWS2.Cells[r, formatCol].Value != null) {
                                //formatString = thisWS2.Cells[r, formatCol].Value;
                                formatString = thisWS2.Cells[r, formatCol].Text.ToString();
                            }

                            //formatCol
                            if (formatString == "z") //Collection automatically say example
                            {
                                thisWS2.Cells[r, c].Value2 = "W";

                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Derived Example set to True because this lot is a collection");
                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Derived Example set to W because this lot is a collection");

                            }
                            else if (nbrInternetPhotos > nbrCatalogPhotos)
                            {
                                thisWS2.Cells[r, c].Value2 = "Y";

                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Derived Example set to True because this lot has more internet photos than catalog photos");
                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Derived Example set to Y because this lot has more internet photos than catalog photos");

                            }
                            else
                            {
                                thisWS2.Cells[r, c].Value2 = "N";

                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("Derived Example set to False because its not a collection and internet photos not > catalog photos");
                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("Derived Example set to N because its not a collection and internet photos not > catalog photos");

                            }
                        }
                        //$$$$$$$
                        //int catalog1Col = 0;
                        //int collectionTypeCol = 0;
                        //int alphaTextCol = 0;
                        //for srtorder, use catalog value if not collection - leave blank if collection
                        else if (thisWS2.Cells[1, c].Value == "SrtOrder") 
                        {
                            if (thisWS2.Cells[r, catalog1Col].Value2 != null && thisWS2.Cells[r, collectionTypeCol].Value2 != "z")
                            {
                                thisWS2.Cells[r, c].Value2 = thisWS2.Cells[r, catalog1Col].Value2;

                                //thisWS2.Cells[r, c].ClearComments();
                                //thisWS2.Cells[r, c].AddComment("SrtOrder set to Catalog 1 Value because catalog 1 number has been assigned and this is not a collection type (z) lot");
                                //thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                thisWS2.Cells[r, c].AddComment("SrtOrder set to Catalog 1 Value because catalog 1 number has been assigned and this is not a collection type (z) lot");
                            }
                            //else
                            //{
                            //    if (thisWS2.Cells[r, alphaTextCol].Value != null)
                            //    {
                            //        thisWS2.Cells[r, c].Value = thisWS2.Cells[r, alphaTextCol].Value;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("SrtOrder set to Alpha Text because catalog 1 number was not present and Alpha Text was present");
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //    }
                            //}
                        }
                        //$$$$$$$$$$$$
                        // https://drive.google.com/drive/folders/1grl8P1eV5HUsjd0_LlLPVHJfh9eukaPz
                        // use the target name for test (i.e. CM.Stamp Symbol maps to SA.Condition)
                        //TODO: HOW TO TRY CATCH THESE IN CASE BAD EXCEL DATA!!!!!!!!!!!!!!!!!!!
                        else if (thisWS2.Cells[1, c].Value2 == "Symbol" && thisWS2.Cells[r, originalSymbolCol].Value2 != null)
                            {
                                string symbols = thisWS2.Cells[r, originalSymbolCol].Value2;
                                var imgs = new List<string>();


                                int i = 0;
                                while ((i = symbols.IndexOf("img", i)) != -1)
                                {
                                    // Print out the substring.
                                    //Console.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);
                                    //Debug.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i, 3), i);

                                    //<img src="http://www.kelleherauctions.com/images/mint.gif" align="top">
                                    string imgname = "";        //TODO: SHOULD DEFINE THESE OUTSIDE OF WHILE?
                                    int start = i + 45;
                                    int backslash = symbols.IndexOf("/", start);
                                    backslash += 1;
                                    int period = symbols.IndexOf(".", start);
                                    int displace = 0;

                                    if (period != -1)
                                    {
                                        displace = period - backslash;

                                        imgname = symbols.Substring(backslash, displace);
                                        //Debug.WriteLine("row: {0} image name {1}", r, imgname);

                                        imgs.Add(imgname);
                                    }

                                    // Increment the index.
                                    i++;
                                }


                                String term = string.Empty;
                                bool inbracket = false;

                                foreach (char ch in symbols)
                                {
                                    if (ch == '<')
                                    {
                                        inbracket = true;
                                        if (term != "")
                                        {
                                            imgs.Add(term);
                                            term = "";
                                        }

                                    }
                                    else
                                        if (ch == '>')
                                    {
                                        inbracket = false;
                                    }
                                    else
                                        if (!inbracket && ch != '/' && ch != '\'' && ch != '\"')
                                    {
                                        term += ch;
                                    }


                                }

                                if (term != "")
                                {
                                    imgs.Add(term);
                                    term = "";
                                }

                                //TODO: SHOULD DEFINE THESE OUTSIDE OF WHILE?
                                string imgid = String.Empty;
                                string origimg = String.Empty;

                                string newSymbols = String.Empty;
                                string newComment = String.Empty;
                                bool wasError = false;

                                if (imgs.Count > 0)
                                {

                                origimg = thisWS2.Cells[r, originalSymbolCol].Value2;
                                ((Excel.Range)thisWS2.Cells[r, c]).Value = "";  //clear out value
                                ((Excel.Range)thisWS2.Cells[r, c]).ClearComments();
                                newComment = "Original symbol before xform: " + origimg;

                                ((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Blue;



                                //%%%%TEST SECTION BEGIN
                                //if (s == "&#41" || s == "&#41;" || s == "&#40" || s == "&#40;")
                                //foreach (string s in imgs)
                                //{
                                //    string s2 = s.Replace("&#41", "");
                                //    s2 = s2.Replace("&#41;", "");
                                //    s2 = s2.Replace("&#40", "");
                                //    s2 = s2.Replace("&#40;", "");
                                //    Debug.WriteLine(s2);
                                //}
                                //%%%%TEST SECTION END


                                //%%%%BEGIN REAL SECTION
                                foreach (string s in imgs)
                                    {

                                        /*
                                        Number sign &#36; 
                                        Dollar sign &#37; 
                                        Percent sign &#38; 
                                        Ampersand &#39; 
                                        Apostrophe &#40; 
                                        Left parenthesis &#41; *******
                                        Right parenthesis &#42; 
                                        Asterisk &#43;
                                        */

                                        if (s == "&#41" || s == "&#41;" || s == "&#40" || s == "&#40;")
                                        {
                                            continue;
                                        }

                                        string s2 = s.Replace("&#41", "");
                                        s2 = s2.Replace("&#41;", "");
                                        s2 = s2.Replace("&#40", "");
                                        s2 = s2.Replace("&#40;", "");

                                        cmd2.CommandText = "select amsid from symbol_reference where cmid = '" + s2 + "'";    //mapping step stuffed CM value, so now re-map
                                        reader2 = cmd2.ExecuteReader();


                                        if (reader2.HasRows)
                                        {
                                            while (reader2.Read())
                                            {

                                                //CMCategoryTxt = reader2.GetString(1);
                                                //SACategoryTxt = reader2.GetString(3);
                                                //EBCategoryTxt = reader2.GetString(5);

                                                imgid = reader2.GetString(0);              //Note EBay id is string until further known
                                                newSymbols = newSymbols + imgid + ",";
                                            //thisWS2.Cells[r, c].ClearComments();
                                            //thisWS2.Cells[r, c].Interior.Color = Color.BurlyWood;
                                            //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.BurlyWood;
                                            //((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuckfuck");
                                            //thisWS2.Cells[r, c].AddComment("fuckfuck");

                                            //thisWS2.Cells[r, c].Value = thisWS2.Cells[r, c].Value + imgid + ";";
                                            //thisWS2.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                            //break;
                                        }
                                    }
                                        else
                                        {
                                        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                                        //thisWS2.Cells[r, c].Interior.Color = Color.Red;

                                        //TODO: add row,column and heading to comment
                                        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                                        //thisWS2.Cells[r, c].ClearComments();
                                        wasError = true;
                                        newComment = newComment + "Tried to find image for CMid " + s2 + " in table Symbol_Reference. CMid not found - please add to table";
                                        //thisWS2.Cells[r, c].AddComment("Tried to find image for CMid " + s2 + " in table Symbol_Reference. CMid not found - please add to table");
                                            //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                                        }

                                        reader2.Close();
                                    }
                                
                                if (newSymbols.Length > 0)
                                {
                                    newSymbols = newSymbols.TrimEnd(',');
                                    //thisWS2.Cells[r, c].Value = newSymbols;
                                    ((Excel.Range)thisWS2.Cells[r, c]).Value2 = newSymbols;
                                    //thisRange.Cells[r, c].Value2 = "bbb";
                                    //thisRange.Cells[r, c].Interior.Color = Color.Blue;

                                    //((Excel.Range)thisWS2.Cells[r, c]).Value = newSymbols;
                                    //((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuck");
                                }

                                //thisWS2.Cells[r, c].AddComment(newComment); //--> OLD WAY STOPPED WORKING
                                //((Excel.Range)fromXlWorksheet.Cells[1, i]).Value --> PROTOTYPE
                                //((Excel.Range)thisWS2.Cells[r, c]).AddComment("fuck");
                                if (wasError)
                                {
                                    thisWS2.Cells[r, c].Interior.Color = Color.Red;
                                }
                                else
                                {
                                    thisRange.Cells[r, c].Interior.Color = Color.Blue;

                                }

                                //%%%%%%%%%%%%%% set value here 
                                //thisWS2.Cells[r, c].AddComment("Original symbol before xform: " + origimg);
                                //%%%%%%%%%%%

                                //%%%%%END REAL SECTION
                            }
                            //***************************************************************************************************************************************************
                            //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                            //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                            //((Excel.Range)thisWS2.Cells[r, reqSaColNbr[c]]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                            //if (thisWS2.Cells[1, c].Value == "AuctionID")
                            //{
                            //    SAAuctionId = 0;

                            //    cmd2.CommandText = "SELECT SAId FROM dbo.Auction where CMId = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            SAAuctionId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SAAuctionId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM sale id: " + thisWS2.Cells[r, c].Value + " to SA auction id - CM sale id not found in Auction table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "CategoryId")
                            //{
                            //    SACategoryId = 0;

                            //    cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            string val = thisWS2.Cells[r, c].Value;

                            //            SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SACategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                            //            thisWS2.Cells[r, c].ClearComments();
                            //            thisWS2.Cells[r, c].AddComment("Mapped CM Country Lookup: " + val + " to SA Category SAID " + SACategoryId);

                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM Country Lookup: " + thisWS2.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "CategoryId2")
                            //{
                            //    SACategoryId = 0;

                            //    cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            string val = thisWS2.Cells[r, c].Value;

                            //            SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SACategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                            //            thisWS2.Cells[r, c].ClearComments();
                            //            thisWS2.Cells[r, c].AddComment("Mapped CM Province Lookup: " + val + " to SA Category SAID " + SACategoryId);

                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM Province Lookup: " + thisWS2.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}

                            //else if (thisWS2.Cells[1, c].Value == "ShippingCategoryId")
                            //{
                            //    SAShippingCategoryId = 0;

                            //    cmd2.CommandText = "SELECT SAid FROM dbo.Shipping_Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {
                            //            string val = thisWS2.Cells[r, c].Value;

                            //            SAShippingCategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SAShippingCategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;

                            //            thisWS2.Cells[r, c].ClearComments();
                            //            thisWS2.Cells[r, c].AddComment("Mapped CM Package Type: " + val + " to SA Shipping_Category (SAID) " + SAShippingCategoryId);

                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM Package Type: " + thisWS2.Cells[r, c].Value + " to SA Shipping_Category id - CM CMCategoryTxt not found in Shipping_Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}

                            //else if (thisWS2.Cells[1, c].Value == "PriceGuide1Id")
                            //{
                            //    SAPriceGuide1Id = 0;

                            //    cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            SAPriceGuide1Id = reader2.GetInt32(0);         //assume it's not red already because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SAPriceGuide1Id;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM PriceGuide1 id: " + thisWS2.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "PriceGuide2Id")
                            //{
                            //    SAPriceGuide2Id = 0;

                            //    cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            SAPriceGuide2Id = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SAPriceGuide2Id;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM PriceGuide2 id: " + thisWS2.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "LOAProviderId")
                            //{
                            //    SALOAProviderId = 0;

                            //    cmd2.CommandText = "SELECT SAId FROM dbo.LOA_Provider where CM_Name = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            SALOAProviderId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SALOAProviderId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //            //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM Certificate Id: " + thisWS2.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "EbayPrimaryCategoryId" && EBayImplemented)
                            //{
                            //    EbayCategoryId = "";

                            //    cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                            //            thisWS2.Cells[r, c].Value = EbayCategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                            //                                                                //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM category: " + thisWS2.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}

                            //%%%%%%%%%%%%%%%%%%%%%%%%%
                            //// https://drive.google.com/drive/folders/1grl8P1eV5HUsjd0_LlLPVHJfh9eukaPz
                            //// use the target name for test (i.e. CM.Stamp Symbol maps to SA.Condition)
                            //else if (thisWS2.Cells[1, c].Value == "Condition")
                            //{
                            //    string symbols = thisWS2.Cells[r, c].Value;
                            //    var imgs = new List<string>();


                            //    int i = 0;
                            //    while ((i = symbols.IndexOf("img", i)) != -1)
                            //    {
                            //        // Print out the substring.
                            //        //Console.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);
                            //        Debug.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i, 3), i);

                            //        //<img src="http://www.kelleherauctions.com/images/mint.gif" align="top">
                            //        string imgname = "";
                            //        int start = i + 45;
                            //        int backslash = symbols.IndexOf("/", start);
                            //        backslash += 1;
                            //        int period = symbols.IndexOf(".", start);
                            //        int displace = 0;

                            //        if (period != -1)
                            //        {
                            //            displace = period - backslash;

                            //            imgname = symbols.Substring(backslash, displace);
                            //            Debug.WriteLine("row: {0} image name {1}", r, imgname);

                            //            imgs.Add(imgname);
                            //        }

                            //        // Increment the index.
                            //        i++;
                            //    }


                            //    String term = string.Empty;
                            //    bool inbracket = false;

                            //    foreach (char ch in symbols)
                            //    {
                            //        if (ch == '<')
                            //        {
                            //            inbracket = true;
                            //            if (term != "")
                            //            {
                            //                imgs.Add(term);
                            //                term = "";
                            //            }

                            //        }
                            //        else
                            //            if (ch == '>')
                            //        {
                            //            inbracket = false;
                            //        }
                            //        else
                            //            if (!inbracket && ch != '/')
                            //        {
                            //            term += ch;
                            //        }


                            //    }

                            //    if (term != "")
                            //    {
                            //        imgs.Add(term);
                            //        term = "";
                            //    }

                            //    string imgid = "";
                            //    string origimg = "";
                            //    if (imgs.Count > 0)
                            //    {

                            //        origimg = thisWS2.Cells[r, c].Value;
                            //        thisWS2.Cells[r, c].Value = ""; //clear out value
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Original condition before xform: " + origimg);


                            //        foreach (string s in imgs)
                            //        {

                            //            /*
                            //            Number sign &#36; 
                            //            Dollar sign &#37; 
                            //            Percent sign &#38; 
                            //            Ampersand &#39; 
                            //            Apostrophe &#40; 
                            //            Left parenthesis &#41; *******
                            //            Right parenthesis &#42; 
                            //            Asterisk &#43;
                            //            */

                            //            if (s == "&#41" || s == "&#41;" || s == "&#40" || s == "&#40;")
                            //            {
                            //                continue;
                            //            }

                            //            cmd2.CommandText = "select id from symbol_reference where cmid = '" + s + "'";    //mapping step stuffed CM value, so now re-map
                            //            reader2 = cmd2.ExecuteReader();


                            //            if (reader2.HasRows)
                            //            {
                            //                while (reader2.Read())
                            //                {

                            //                    //CMCategoryTxt = reader2.GetString(1);
                            //                    //SACategoryTxt = reader2.GetString(3);
                            //                    //EBCategoryTxt = reader2.GetString(5);

                            //                    imgid = reader2.GetString(0);              //Note EBay id is string until further known
                            //                    thisWS2.Cells[r, c].Value = thisWS2.Cells[r, c].Value + imgid + ";";
                            //                    thisWS2.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                            //                                                                        //break;
                            //                }
                            //            }
                            //            else
                            //            {
                            //                //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //                thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //                //TODO: add row,column and heading to comment
                            //                //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //                thisWS2.Cells[r, c].ClearComments();
                            //                thisWS2.Cells[r, c].AddComment("Tried to find image for CMid " + s + " in table Symbol_Reference. CMid not found");
                            //                //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //            }

                            //            reader2.Close();
                            //        }
                            //    }

                            //    //string sub = input.Substring(0, 3);
                            //    //Console.WriteLine("Substring: {0}", sub);

                            //}

                            //else if (thisWS2.Cells[1, c].Value == "EbaySecondaryCategoryId" && EBayImplemented)
                            //{
                            //    EbayCategoryId = "";

                            //    cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                            //            thisWS2.Cells[r, c].Value = EbayCategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                            //                                                                //break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].ClearComments();
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM category: " + thisWS2.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //    }

                            //    reader2.Close();
                            //}
                            //else if (thisWS2.Cells[1, c].Value == "SerialNumber")
                            //{

                            //    if (thisWS2.Cells[r, c].Value != null && thisWS2.Cells[r, c].Value != 0)
                            //    {

                            //        //var client = new RestClient("http://kelleherdemo2-com.si-sv2521.com");
                            //        var client = new RestClient("http://kelleher-stage-com.si-sv2521.com/Kelleher.aspx?debug=GetInventoryIdBySerialNumber");

                            //        var request = new RestRequest("/Kelleher.aspx", Method.POST);
                            //        request.RequestFormat = DataFormat.Json;

                            //        double serialNbrStr = thisWS2.Cells[r, c].Value;
                            //        //string serialNbrStr = thisWS2.Cells[r, c].Value;
                            //        //string serialNbrStr = serialNbr.ToString();

                            //        //https://stackoverflow.com/questions/14828520/how-to-create-my-json-string-by-using-c
                            //        //var f = new SARestLoginModel
                            //        //{
                            //        //    request = new Dictionary<string, string>
                            //        //    {
                            //        //        {"username", "admin"},
                            //        //        {"password", "admin"},
                            //        //        {"operation", "GetConsignors"},
                            //        //        //{"serialnumber", serialNbrStr},
                            //        //    }
                            //        //};

                            //        //{"request":{"username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}}
                            //        //request.AddJsonBody(new { A = "foo", B = "bar" });
                            //        //request.AddJsonBody(new { "request":{ "username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}
                            //        request.AddJsonBody(new { request = new { username = "admin", password = "admin", operation = "GetInventoryIdBySerialNumber", serialnumber = serialNbrStr.ToString() } });

                            //        //});

                            //        //request.AddBody(f);
                            //        //request.AddXmlBody(f);

                            //        IRestResponse response = client.Execute(request);

                            //        if (!response.IsSuccessful)
                            //        {
                            //            Debug.WriteLine("reponse failed");
                            //        }

                            //        JObject obj1 = JObject.Parse(response.Content);
                            //        //JArray SAInventoryId = (JArray)obj1["inventoryid"];
                            //        JValue SAInventoryId = (JValue)obj1["inventoryid"];

                            //        //int len = SAInventoryId.Count;
                            //        //int inventoryId = 0;
                            //        string inventoryId = (string)SAInventoryId;
                            //        //inventoryId = (int)SAInventoryId[0]["id"]; //????


                            //        if (inventoryId != null && inventoryId != "-1")            //TODO: CHECK FOR -1
                            //        {
                            //            //int col = GetInvColumn(thisWS2,r);

                            //            for (int i = 1; i < 256; i++)
                            //            {
                            //                if (thisWS2.Cells[1, i].Value == "InventoryId")
                            //                {
                            //                    //return (i);
                            //                    thisWS2.Cells[r, i].Value = inventoryId; //TODO BREAK WHEN HIT FIRST ONE
                            //                    break;
                            //                }
                            //            }
                            //        }


                            //        //thisWS2.Cells[r, c].Value = EbayCategoryId;
                            //        //thisWS2.Cells[r, c].Interior.Color = Color.Blue;


                            //        //SAInventoryId = 0;

                            //        //cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //        //reader2 = cmd2.ExecuteReader();


                            //        //if (reader2.HasRows)
                            //        //{
                            //        //    while (reader2.Read())
                            //        //    {

                            //        //        //CMCategoryTxt = reader2.GetString(1);
                            //        //        //SACategoryTxt = reader2.GetString(3);
                            //        //        //EBCategoryTxt = reader2.GetString(5);

                            //        //        SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //        //        thisWS2.Cells[r, c].Value = SAConsignorId;
                            //        //        thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //        //        //break;
                            //        //    }
                            //        //}
                            //        //else
                            //        //{
                            //        //    //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        //    thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //    //TODO: add row,column and heading to comment
                            //        //    //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        //    thisWS2.Cells[r, c].ClearComments();
                            //        //    thisWS2.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS2.Cells[r, c].Value + " to SA consignor id - CM consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //        //    //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
                            //        //}

                            //        //reader2.Close();
                            //    }
                            //}






                            //**** Map Consignor Id ****
                            //if (thisWS2.Cells[1, c].Value == "CategoryId")
                            //{
                            //    SACategoryId = 0;

                            //    cmd2.CommandText = "SELECT * FROM dbo.Category where CMCategoryTxt = '" + thisWS2.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            //    reader2 = cmd2.ExecuteReader();


                            //    if (reader2.HasRows)
                            //    {
                            //        while (reader2.Read())
                            //        {

                            //            //CMCategoryTxt = reader2.GetString(1);
                            //            //SACategoryTxt = reader2.GetString(3);
                            //            //EBCategoryTxt = reader2.GetString(5);

                            //            SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                            //            thisWS2.Cells[r, c].Value = SACategoryId;
                            //            thisWS2.Cells[r, c].Interior.Color = Color.Blue;
                            //            break;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
                            //        thisWS2.Cells[r, c].Interior.Color = Color.Red;

                            //        //TODO: add row,column and heading to comment
                            //        //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
                            //        thisWS2.Cells[r, c].AddComment("Tried to map CM category: " + thisWS2.Cells[r, c].Value + " to SA category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                            //    }
                            //    reader2.Close();
                            //}

                        }
                        }
                    }
            }


            ////TODO: THIS ASSUMES YOU'LL ALWAYS STUFF LOT NUMBER, EVEN IF LOTS ALREADY SEQUENCED IN SA. PHASE II SHOULD PULL LOT NUMBERS FOR CHANGE IN CASE THEY WERE SEQUENCED
            ////Stuff Lot Numbers
            //string desc = "";
            ////int lotnbr = 990000;
            //int lotcol = 0;
            //for (int c = 1; c <= colCount; c++)
            //{
            //    if (thisWS2.Cells[1, c].Value == "LotNumber")
            //    {
            //        lotcol = c;
            //        break;
            //    }
            //}

            //int desccol = 0;
            //for (int c = 1; c <= colCount; c++)
            //{
            //    if (thisWS2.Cells[1, c].Value == "Description")
            //    {
            //        desccol = c;
            //        break;
            //    }
            //}

            //if (lotcol != 0 && desccol != 0)
            //{
            //    for (int r = 2; r <= rowCount; r++)
            //    {
            //        //thisWS2.Cells[r, lotcol].Value = lotnbr;
            //        //lotnbr += 1;

            //        desc = "";

            //        cmd2.CommandText = "select descrip from SAN_Sale_Data  where LOT_NO = '" + thisWS2.Cells[r, lotcol].Value + "'";    //mapping step stuffed CM value, so now re-map
            //        reader2 = cmd2.ExecuteReader();

            //        if (reader2.HasRows)
            //        {
            //            while (reader2.Read())
            //            {

            //                //CMCategoryTxt = reader2.GetString(1);
            //                //SACategoryTxt = reader2.GetString(3);
            //                //EBCategoryTxt = reader2.GetString(5);

            //                desc = reader2.GetString(0);              //Note EBay id is string until further known
            //                thisWS2.Cells[r, desccol].Value = desc;
            //                thisWS2.Cells[r, desccol].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup

            //            }
            //        }
            //        else
            //        {
            //            //((Excel.Range)thisWS2.Cells[r, c]).Interior.Color = Color.Red;
            //            thisWS2.Cells[r, desccol].Interior.Color = Color.Red;

            //            //TODO: add row,column and heading to comment
            //            //((Excel.Range)thisWS2.Cells[r, x]).AddComment(thisWS2.Cells[r, c].Value + " is required") ;
            //            thisWS2.Cells[r, desccol].ClearComments();
            //            thisWS2.Cells[r, desccol].AddComment("Tried to find description DESCRIP in the SAN data for this lots: " + thisWS2.Cells[r, lotcol].Value);
            //            //thisWS2.Cells[r, c].Comment[1].AutoFit = true;
            //        }

            //        reader2.Close();
            //    }
            //}


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

        private void btnSelectCategory_Click_1(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;


            if (!thisWS.AutoFilterMode)
            {
                MessageBox.Show("Autofilter must be on to assign category codes ");
                return;
            }

            //MessageBox.Show("da fuck!");
            var form1 = new Form1();
            form1.Show();

        }

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

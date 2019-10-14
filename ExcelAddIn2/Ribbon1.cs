using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Data.SqlClient;
using System.Collections;

using System.Drawing;

using System.Diagnostics;

using RestSharp;
using RestSharp.Authenticators;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
        }

        public const bool EBayImplemented = false;


        //******************************************************************************************************************************************
        //Get reference to current sheet (with add-in) that will hold resulting SA Lot file
        //*****************************************************************************************************************************************


        ArrayList headingsMap = new ArrayList();



        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnLoadAMS_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS.UsedRange;
            int thisRowCount = thisRange.Rows.Count;
            int thisColCount = thisRange.Columns.Count;

            // thisWS.Cells[thisRowCount, thisColCount].Cleasarestr();               //this bullshit doesn't work
            // thisWS.Cells[thisRowCount, thisColCount].ClearComments();        //this bullshit doesn't work

            //** Clear Comments
            for (int r = 1; r <= thisRowCount; r++)
            {
                for (int c = 1; c <= thisColCount; c++)
                {
                    thisWS.Cells[r, c].ClearComments();
                    thisWS.Cells[r, c].Clear();
                }
            }


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

            string msg = "";  //Trace msg

            OneColumnMap thisColumnMap;


            //SqlConnection sqlConnection1 = new SqlConnection("Data Source=BACKUPDELL\\SQLEXPRESS ;Initial Catalog=SimpleAuction;Integrated Security=True");
            //SqlConnection sqlConnection1 = new SqlConnection("Data Source=BACKUPDELL ;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection1 = new SqlConnection("Data Source=MANCINI-AWARE ;Initial Catalog=Describing;Integrated Security=True");
            SqlCommand cmd1 = new SqlCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.Connection = sqlConnection1;
            SqlDataReader reader1;
            cmd1.CommandText = "SELECT SAColumnNbr,SAHeading,CMColumnNbr,CMHeading, Required, DefaultValue, mapDB, SARequired, Note FROM dbo.ExcelHeadingMap order by SAColumnNbr";

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

                    //Populate first row current column with SA heading and comment describing how it is derived
                    ((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
                    if (thisColumnMap.CMHead != null)
                    {
                        thisWS.Cells[1, thisColumnMap.SAPosition].ClearComments();
                        thisWS.Cells[1, thisColumnMap.SAPosition].AddComment("Pulled from Catalog Master field: " + thisColumnMap.CMHead + ".  " + thisColumnMap.Note);
                    }



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


                    msg = "FROM STRUCT --> SAPosition: " + thisColumnMap.SAPosition.ToString() + "/SAHead: " + thisColumnMap.SAHead + "/CMPosition: " + thisColumnMap.CMPosition.ToString() + "/CMHead: " + thisColumnMap.CMHead;
                    Trace.WriteLine(msg + "\t");


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


            //Declare reuseable (per file) variables here
            var fromXlApp = new Excel.Application();
            Excel.Workbook fromXlWorkbook;
            Excel._Worksheet fromXlWorksheet;
            Excel.Range fromXlRange;
            int rowCount;
            int colCount;

            int filecount = 0;
            //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

            foreach (string fn in openFileDialog1.FileNames)
            {


                //TODO: CHECK FILE NAME BEFORE OPENING AGAINST ??? TO MAKE SURE IT'S AN ?UNPROCESSED? ?NEW? ?WELL-NAMED? CATALOG MASTER FILE
                //filename = openFileDialog1.FileName;
                filename = fn;
                MessageBox.Show("For pause - here is file name about to process: " + filename);

                //excelApp.StatusBar = String.Format("Processing line {0} on {1}.",rows,rowNum);
                Globals.ThisAddIn.Application.StatusBar = String.Format("Loading file {0}: {1}", filecount + 1, openFileDialog1.SafeFileNames[filecount]);

                //%%%%%%%%%%%%%%% START LOOP HERE

                //*var fromXlApp = new Excel.Application();
                fromXlApp.Visible = false; //--> Don't need to see the Catalog Master excel file to suck it in

                fromXlWorkbook = fromXlApp.Workbooks.Open(filename);     //this is the fully qualified (local) file name
                fromXlWorksheet = fromXlWorkbook.Sheets[1];            //TODO: make sure only one worksheet???
                fromXlRange = fromXlWorksheet.UsedRange;


                //MessageBox.Show("CM (from) file should be open now ... begin data map/load from CM to current SA lot spreadsheet");
                //Globals.ThisAddIn.Application.StatusBar = "processing file";
                Cursor.Current = Cursors.Hand;

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

                //for (int r = 2; r <= rowCount; r++)         //r = TO ROW TO FILL - WALK HEADING MAP ARRAY TO LEARN COLUMNS TO COPY
                //{
                //SEE IF ANY HEADINGS ARE MAPPED FOR THIS "TO" ROW
                foreach (OneColumnMap map in headingsMap)
                {
                    if (map.SAHead != "" && map.CMHead == "" && map.defaultValue == "")
                    {
                        continue;
                    }
                    else
                        if (map.SAHead != "" && map.CMHead != "")           //if CM heading mapped to SA Heading move the cm spreadsheet value
                    {
                        //use the current row in the "TO" spreadsheet- (outer loop)
                        //NOTE: this will bring mapped fields over as well. MAPPING will occur in ValidateSpreadsheet();
                        //NOTE: This assumes column headings are uniquie within spreadsheets
                        for (int i = 1; i < 256; i++)
                        {
                            if (thisWS.Cells[1, i].Value == map.SAHead)
                            {
                                SAHeadCol = i;
                            }

                            if (fromXlRange.Cells[1, i].Value == map.CMHead)
                            {
                                CMHeadCol = i;
                            }
                        }

                        if (SAHeadCol != 0 && CMHeadCol != 0)       //TODO: LOG IF DB HAS COLUMN MAPPINGS BUT YOU CAN'T FIND COLUMS IN FROM OR TO SPREADSHEET
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value = (fromXlRange.Cells[r, CMHeadCol].Value);
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
                            if (thisWS.Cells[1, i].Value == map.SAHead)
                            {
                                SAHeadCol = i;
                            }
                        }

                        if (SAHeadCol != 0)
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value = map.defaultValue;
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
                //Marshal.ReleaseComObject(fromXlWorksheet);

                //close and release
                fromXlWorkbook.Close();
                //Marshal.ReleaseComObject(fromXlWorkbook);

                //quit and release
                fromXlApp.Quit();
                // Marshal.ReleaseComObject(fromXlApp);

                filecount += 1;

            }
            //%%%% end File Name loop



            //*************************************************************************************************************************************************************
            //* This will check for missing values and also map database (id) values for category, auction/sale, consignor, consignment
            //*************************************************************************************************************************************************************
            ValidateSpreadsheet();

            //%%%%%%%%%set lot numbers

            Globals.ThisAddIn.Application.StatusBar = String.Format("All Catalog Master files are loaded and Validation is complete");

            Cursor.Current = Cursors.Default;   //TODO: HOW TO CLEAN UP ALL EXCEL OBJECTS
        }
        //----------------------------------------------------------------------------------------------------------------------------------

        private void btnLoadCatMast_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS.UsedRange;
            int thisRowCount = thisRange.Rows.Count;
            int thisColCount = thisRange.Columns.Count;

            // thisWS.Cells[thisRowCount, thisColCount].Cleasarestr();               //this bullshit doesn't work
            // thisWS.Cells[thisRowCount, thisColCount].ClearComments();        //this bullshit doesn't work

            //** Clear Comments
            for (int r = 1; r <= thisRowCount; r++)
            {
                for (int c = 1; c <= thisColCount; c++)
                {
                    thisWS.Cells[r, c].ClearComments();
                    thisWS.Cells[r, c].Clear();
                }
            }


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

            string msg = "";  //Trace msg

            OneColumnMap thisColumnMap;


            //SqlConnection sqlConnection1 = new SqlConnection("Data Source=BACKUPDELL\\SQLEXPRESS ;Initial Catalog=SimpleAuction;Integrated Security=True");
            //SqlConnection sqlConnection1 = new SqlConnection("Data Source=BACKUPDELL ;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection1 = new SqlConnection("Data Source=MANCINI-AWARE ;Initial Catalog=Describing;Integrated Security=True");
            SqlCommand cmd1 = new SqlCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.Connection = sqlConnection1;
            SqlDataReader reader1;
            cmd1.CommandText = "SELECT SAColumnNbr,SAHeading,CMColumnNbr,CMHeading, Required, DefaultValue, mapDB, SARequired, Note FROM dbo.ExcelHeadingMap order by SAColumnNbr";

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

                    //Populate first row current column with SA heading and comment describing how it is derived
                    ((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
                    if (thisColumnMap.CMHead != null)
                    {
                        thisWS.Cells[1, thisColumnMap.SAPosition].ClearComments();
                        thisWS.Cells[1, thisColumnMap.SAPosition].AddComment("Pulled from Catalog Master field: " + thisColumnMap.CMHead + ".  " + thisColumnMap.Note);
                    }



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


                    msg = "FROM STRUCT --> SAPosition: " + thisColumnMap.SAPosition.ToString() + "/SAHead: " + thisColumnMap.SAHead + "/CMPosition: " + thisColumnMap.CMPosition.ToString() + "/CMHead: " + thisColumnMap.CMHead;
                    Trace.WriteLine(msg + "\t");


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


            //Declare reuseable (per file) variables here
            var fromXlApp = new Excel.Application();
            Excel.Workbook fromXlWorkbook;
            Excel._Worksheet fromXlWorksheet;
            Excel.Range fromXlRange;
            int rowCount;
            int colCount;

            int filecount = 0;
            //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

            foreach (string fn in openFileDialog1.FileNames)
            {


                //TODO: CHECK FILE NAME BEFORE OPENING AGAINST ??? TO MAKE SURE IT'S AN ?UNPROCESSED? ?NEW? ?WELL-NAMED? CATALOG MASTER FILE
                //filename = openFileDialog1.FileName;
                filename = fn;
                MessageBox.Show("For pause - here is file name about to process: " + filename);

                //excelApp.StatusBar = String.Format("Processing line {0} on {1}.",rows,rowNum);
                Globals.ThisAddIn.Application.StatusBar = String.Format("Loading file {0}: {1}", filecount + 1, openFileDialog1.SafeFileNames[filecount]);

                //%%%%%%%%%%%%%%% START LOOP HERE

                //*var fromXlApp = new Excel.Application();
                fromXlApp.Visible = false; //--> Don't need to see the Catalog Master excel file to suck it in

                fromXlWorkbook = fromXlApp.Workbooks.Open(filename);     //this is the fully qualified (local) file name
                fromXlWorksheet = fromXlWorkbook.Sheets[1];            //TODO: make sure only one worksheet???
                fromXlRange = fromXlWorksheet.UsedRange;


                //MessageBox.Show("CM (from) file should be open now ... begin data map/load from CM to current SA lot spreadsheet");
                //Globals.ThisAddIn.Application.StatusBar = "processing file";
                Cursor.Current = Cursors.Hand;

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

                //for (int r = 2; r <= rowCount; r++)         //r = TO ROW TO FILL - WALK HEADING MAP ARRAY TO LEARN COLUMNS TO COPY
                //{
                //SEE IF ANY HEADINGS ARE MAPPED FOR THIS "TO" ROW
                foreach (OneColumnMap map in headingsMap)
                {
                    if (map.SAHead != "" && map.CMHead == "" && map.defaultValue == "")
                    {
                        continue;
                    }
                    else
                        if (map.SAHead != "" && map.CMHead != "")           //if CM heading mapped to SA Heading move the cm spreadsheet value
                    {
                        //use the current row in the "TO" spreadsheet- (outer loop)
                        //NOTE: this will bring mapped fields over as well. MAPPING will occur in ValidateSpreadsheet();
                        //NOTE: This assumes column headings are uniquie within spreadsheets
                        for (int i = 1; i < 256; i++)
                        {
                            if (thisWS.Cells[1, i].Value == map.SAHead)
                            {
                                SAHeadCol = i;
                            }

                            if (fromXlRange.Cells[1, i].Value == map.CMHead)
                            {
                                CMHeadCol = i;
                            }
                        }

                        if (SAHeadCol != 0 && CMHeadCol != 0)       //TODO: LOG IF DB HAS COLUMN MAPPINGS BUT YOU CAN'T FIND COLUMS IN FROM OR TO SPREADSHEET
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value = (fromXlRange.Cells[r, CMHeadCol].Value);
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
                            if (thisWS.Cells[1, i].Value == map.SAHead)
                            {
                                SAHeadCol = i;
                            }
                        }

                        if (SAHeadCol != 0)
                        {
                            for (int r = 2; r <= rowCount; r++)
                            {
                                ((Excel.Range)thisWS.Cells[r, SAHeadCol]).Value = map.defaultValue;
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
                //Marshal.ReleaseComObject(fromXlWorksheet);

                //close and release
                fromXlWorkbook.Close();
                //Marshal.ReleaseComObject(fromXlWorkbook);

                //quit and release
                fromXlApp.Quit();
                // Marshal.ReleaseComObject(fromXlApp);

                filecount += 1;

            }
            //%%%% end File Name loop



            //*************************************************************************************************************************************************************
            //* This will check for missing values and also map database (id) values for category, auction/sale, consignor, consignment
            //*************************************************************************************************************************************************************
            ValidateSpreadsheet();

            //%%%%%%%%%set lot numbers

            Globals.ThisAddIn.Application.StatusBar = String.Format("All Catalog Master files are loaded and Validation is complete");

            Cursor.Current = Cursors.Default;   //TODO: HOW TO CLEAN UP ALL EXCEL OBJECTS
        }


        //git change again
        
        private void ValidateSpreadsheet()
        {
            //((Excel.Range) thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            Globals.ThisAddIn.Application.StatusBar = String.Format("All files loaded - Now we are validating the target SimpleAuction spreadsheet. Please be patient.");

            Cursor.Current = Cursors.WaitCursor;

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            
            //*****************************************************************************************************************************************************
            //* Load up the database driven mapped fields: categor, sale, consignor, consignment
            //****************************************************************************************************************************************************


            //TODO: protect ranges where you don't want them to change directly


            //Set up connection for multiple queries
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL\\SQLEXPRESS ;Initial Catalog=SimpleAuction;Integrated Security=True");
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection2 = new SqlConnection("Data Source=MANCINI-AWARE;Initial Catalog=Describing;Integrated Security=True");
            SqlCommand cmd2 = new SqlCommand();
            cmd2.CommandType = CommandType.Text;
            cmd2.Connection = sqlConnection2;
            sqlConnection2.Open();

            SqlDataReader reader2;


            
            //TODO: MAKE SURE GETTING LAST ROW AND LAST COLUMN

            //****************************************************************************************************************************************************************************
            //* Validate Required fields (indicated in ExcelHeadingMap
            //****************************************************************************************************************************************************************************
            Excel.Range thisRange = thisWS.UsedRange;
            int rowCount = thisRange.Rows.Count;
            int colCount = thisRange.Columns.Count;

            //build list of SA columns that are required
            //ArrayList reqSaColNbr = new ArrayList();
            ArrayList reqSaColName = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.Required)           //if CM heading mapped to SA Heading
                {
                    //reqSaColNbr.Add(map.SAPosition);
                    reqSaColName.Add(map.SAHead);
                 
                }
            }

            //get count of required columns to walk
            int requiredCount = reqSaColName.Count;
            //int requiredCount = reqSaColNbr.Count;

            //build list of SA columns that are mapped

            //ArrayList mapSaColNbr = new ArrayList();
            ArrayList mapSaColName = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.mapDB)           //if CM heading mapped to SA Heading
                {
                    //mapSaColNbr.Add(map.SAPosition);
                    mapSaColName.Add(map.SAHead);
                }
            }




            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //use the current row in the "TO" spreadsheet- (outer loop)
            //((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Interior.Color = Color.White;
            // ((Excel.Range)thisWS.Cells[rowCount, colCount]).ClearComments();
            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Clear();   //SHOULD RESET COLOR AND COMMENTS INSTEAD OF ABOVE


            //*********************************************************************************************************************************************************
            //* Edit for required fields first
            //*********************************************************************************************************************************************************
            //System.Type type;

            for (int r = 2; r <= rowCount; r++) {
                for (int c = 1; c <= colCount; c++) {
                    //value = thisWS.Cells[r, c].Value;
                    //valueint = 0;
                    //o = Int32.TryParse(value, out valueint);

                  
                    foreach (string reqSAName in reqSaColName) {
                        //value = thisWS.Cells[r, c].Value;
                        //if ((c == reqSACol) && ((thisWS.Cells[r, c].Value == null) || (valueint == 0))) {

                        //TODO --> type = (thisWS.Cells[r, c].Value).GetType;  //TODO: NEED TO CHECK FOR 0 IN SOME FIELDS I.E. SEQUENCE NBR
                        if (thisWS.Cells[1, c].Value == reqSAName && thisWS.Cells[r, c].Value == null) {
                            ((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;

                            //TODO: add row,column and heading to comment
                            String txt = thisWS.Cells[1, c].Value;
                            thisWS.Cells[r, c].ClearComments();
                            thisWS.Cells[r, c].AddComment( txt + " is required") ;
                            //((Excel.Range)ws.Cells[r, c]).Style.Name = "Normal"
                        }
                    }
                }
            }

            //****************************************************************************************************************************************************************************
            // Field Mapping exercises - IS Field title specific - no longer relies on column number
            //****************************************************************************************************************************************************************************

            int SACategoryId = 0;
            string EbayCategoryId = "";
            int SAAuctionId = 0;
            int SAConsignorId = 0;
            int SAConsignmentId = 0;
            int SAPriceGuide1Id = 0;
            int SAPriceGuide2Id = 0;
            int SALOAProviderId = 0;
            int SAShippingCategoryId = 0;

            //Check fields that require database mapping by ID
            for (int r = 2; r <= rowCount; r++) {
                for (int c = 1; c <= colCount; c++) {
                        if ((thisWS.Cells[1, c].Value != null) && (thisWS.Cells[r, c].Value != null))   //mapped columns are required so will be red if not provided
                        {
                        //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                        //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                        //((Excel.Range)thisWS.Cells[r, reqSaColNbr[c]]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                        if (thisWS.Cells[1, c].Value == "AuctionID")
                        {
                            SAAuctionId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Auction where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAAuctionId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAAuctionId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM sale id: " + thisWS.Cells[r, c].Value + " to SA auction id - CM sale id not found in Auction table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "CategoryId")
                        {
                            SACategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    string val = thisWS.Cells[r, c].Value;

                                    SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SACategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Country Lookup: " + val + " to SA Category SAID " + SACategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Country Lookup: " + thisWS.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "CategoryId2")
                        {
                            SACategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    string val = thisWS.Cells[r, c].Value;

                                    SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SACategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Province Lookup: " + val + " to SA Category SAID " + SACategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Province Lookup: " + thisWS.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "ConsignerId")
                        {
                            SAConsignorId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAConsignorId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS.Cells[r, c].Value + " to SA consignor id - CM consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        //****************
                        else if (thisWS.Cells[1, c].Value == "ConsignmentTrackingId")
                        {
                            SAConsignmentId = 0;
                            

                            cmd2.CommandText = "SELECT SAId FROM dbo.Consignment where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAConsignmentId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAConsignmentId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM consignment id: " + thisWS.Cells[r, c].Value + " to SA consignment id - CM consignment id not found in Consignment table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "ShippingCategoryId")
                        {
                            SAShippingCategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Shipping_Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {
                                    string val = thisWS.Cells[r, c].Value;

                                    SAShippingCategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAShippingCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Package Type: " + val + " to SA Shipping_Category (SAID) " + SAShippingCategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Package Type: " + thisWS.Cells[r, c].Value + " to SA Shipping_Category id - CM CMCategoryTxt not found in Shipping_Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }

                        else if (thisWS.Cells[1, c].Value == "PriceGuide1Id")
                        {
                            SAPriceGuide1Id = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAPriceGuide1Id = reader2.GetInt32(0);         //assume it's not red already because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAPriceGuide1Id;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM PriceGuide1 id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "PriceGuide2Id")
                        {
                            SAPriceGuide2Id = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAPriceGuide2Id = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAPriceGuide2Id;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM PriceGuide2 id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "LOAProviderId")
                        {
                            SALOAProviderId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.LOA_Provider where CM_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SALOAProviderId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SALOAProviderId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Certificate Id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "EbayPrimaryCategoryId" && EBayImplemented)
                        {
                            EbayCategoryId = "";

                            cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                                    thisWS.Cells[r, c].Value = EbayCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                        //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }

                        // https://drive.google.com/drive/folders/1grl8P1eV5HUsjd0_LlLPVHJfh9eukaPz
                        else if (thisWS.Cells[1, c].Value == "Condition")
                        {
                            string symbols = thisWS.Cells[r, c].Value;
                            var imgs = new List<string>();
               

                            int i = 0;
                            while ((i = symbols.IndexOf("img", i)) != -1)
                            {
                                // Print out the substring.
                                //Console.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);
                                Debug.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);

                                //<img src="http://www.kelleherauctions.com/images/mint.gif" align="top">
                                string imgname = "";
                                int start = i + 45;
                                int backslash = symbols.IndexOf("/", start);
                                backslash += 1;
                                int period = symbols.IndexOf(".", start);
                                int displace = 0;

                                if (period != -1)
                                {
                                    displace = period - backslash;

                                    imgname = symbols.Substring(backslash, displace);
                                    Debug.WriteLine("row: {0} image name {1}", r, imgname);

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
                                    if (!inbracket && ch != '/')
                                {
                                    term += ch;
                                }


                            }

                            if (term != "")
                            {
                                imgs.Add(term);
                                term = "";
                            }

                            string imgid = "";
                            string origimg = "";
                            if (imgs.Count > 0){

                                origimg = thisWS.Cells[r, c].Value;
                                thisWS.Cells[r, c].Value = ""; //clear out value
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Original condition before xform: " + origimg);


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

                                    cmd2.CommandText = "select id from symbol_reference where cmid = '" + s + "'";    //mapping step stuffed CM value, so now re-map
                                    reader2 = cmd2.ExecuteReader();


                                    if (reader2.HasRows)
                                    {
                                        while (reader2.Read())
                                        {

                                            //CMCategoryTxt = reader2.GetString(1);
                                            //SACategoryTxt = reader2.GetString(3);
                                            //EBCategoryTxt = reader2.GetString(5);

                                            imgid = reader2.GetString(0);              //Note EBay id is string until further known
                                            thisWS.Cells[r, c].Value = thisWS.Cells[r, c].Value + imgid + ";";
                                            thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                                //break;
                                        }
                                    }
                                    else
                                    {
                                        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                        thisWS.Cells[r, c].Interior.Color = Color.Red;

                                        //TODO: add row,column and heading to comment
                                        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                        thisWS.Cells[r, c].ClearComments();
                                        thisWS.Cells[r, c].AddComment("Tried to find image for CMid " +  s + " in table Symbol_Reference. CMid not found");
                                        //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                    }

                                    reader2.Close();
                                }
                            }

                                //string sub = input.Substring(0, 3);
                                //Console.WriteLine("Substring: {0}", sub);

                        }

                        else if (thisWS.Cells[1, c].Value == "EbaySecondaryCategoryId" && EBayImplemented)
                        {
                            EbayCategoryId = "";

                            cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                                    thisWS.Cells[r, c].Value = EbayCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                        //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "SerialNumber")
                        {

                            if (thisWS.Cells[r, c].Value != null && thisWS.Cells[r, c].Value != 0)
                            {

                                //var client = new RestClient("http://kelleherdemo2-com.si-sv2521.com");
                                var client = new RestClient("http://kelleher-stage-com.si-sv2521.com/Kelleher.aspx?debug=GetInventoryIdBySerialNumber");

                                var request = new RestRequest("/Kelleher.aspx", Method.POST);
                                request.RequestFormat = DataFormat.Json;

                                double serialNbrStr = thisWS.Cells[r, c].Value;
                                //string serialNbrStr = thisWS.Cells[r, c].Value;
                                //string serialNbrStr = serialNbr.ToString();

                                //https://stackoverflow.com/questions/14828520/how-to-create-my-json-string-by-using-c
                                //var f = new SARestLoginModel
                                //{
                                //    request = new Dictionary<string, string>
                                //    {
                                //        {"username", "admin"},
                                //        {"password", "admin"},
                                //        {"operation", "GetConsignors"},
                                //        //{"serialnumber", serialNbrStr},
                                //    }
                                //};

                                //{"request":{"username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}}
                                //request.AddJsonBody(new { A = "foo", B = "bar" });
                                //request.AddJsonBody(new { "request":{ "username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}
                                request.AddJsonBody(new { request = new { username = "admin", password = "admin", operation = "GetInventoryIdBySerialNumber", serialnumber = serialNbrStr.ToString() } });

                                //});

                                //request.AddBody(f);
                                //request.AddXmlBody(f);

                                IRestResponse response = client.Execute(request);

                                if (!response.IsSuccessful)
                                {
                                    Debug.WriteLine("reponse failed");
                                }

                                JObject obj1 = JObject.Parse(response.Content);
                                //JArray SAInventoryId = (JArray)obj1["inventoryid"];
                                JValue SAInventoryId = (JValue)obj1["inventoryid"];

                                //int len = SAInventoryId.Count;
                                //int inventoryId = 0;
                                string inventoryId = (string)SAInventoryId;
                                //inventoryId = (int)SAInventoryId[0]["id"]; //????


                                if (inventoryId != null && inventoryId != "-1")            //TODO: CHECK FOR -1
                                {
                                    //int col = GetInvColumn(thisWS,r);

                                    for (int i = 1; i < 256; i++)
                                    {
                                        if (thisWS.Cells[1, i].Value == "InventoryId")
                                        {
                                            //return (i);
                                            thisWS.Cells[r, i].Value = inventoryId; //TODO BREAK WHEN HIT FIRST ONE
                                            break;
                                        }
                                    }
                                }


                                //thisWS.Cells[r, c].Value = EbayCategoryId;
                                //thisWS.Cells[r, c].Interior.Color = Color.Blue;


                                //SAInventoryId = 0;

                                //cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                                //reader2 = cmd2.ExecuteReader();


                                //if (reader2.HasRows)
                                //{
                                //    while (reader2.Read())
                                //    {

                                //        //CMCategoryTxt = reader2.GetString(1);
                                //        //SACategoryTxt = reader2.GetString(3);
                                //        //EBCategoryTxt = reader2.GetString(5);

                                //        SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                //        thisWS.Cells[r, c].Value = SAConsignorId;
                                //        thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                //        //break;
                                //    }
                                //}
                                //else
                                //{
                                //    //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                //    thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //    //TODO: add row,column and heading to comment
                                //    //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                //    thisWS.Cells[r, c].ClearComments();
                                //    thisWS.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS.Cells[r, c].Value + " to SA consignor id - CM consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //    //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                //}

                                //reader2.Close();
                            }
                        }

                        




                                //**** Map Consignor Id ****
                                //if (thisWS.Cells[1, c].Value == "CategoryId")
                                //{
                                //    SACategoryId = 0;

                                //    cmd2.CommandText = "SELECT * FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                                //    reader2 = cmd2.ExecuteReader();


                                //    if (reader2.HasRows)
                                //    {
                                //        while (reader2.Read())
                                //        {

                                //            //CMCategoryTxt = reader2.GetString(1);
                                //            //SACategoryTxt = reader2.GetString(3);
                                //            //EBCategoryTxt = reader2.GetString(5);

                                //            SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                //            thisWS.Cells[r, c].Value = SACategoryId;
                                //            thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                //            break;
                                //        }
                                //    }
                                //    else
                                //    {
                                //        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                //        thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //        //TODO: add row,column and heading to comment
                                //        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                //        thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to SA category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //    }



                                //    reader2.Close();
                                //}



                            }
                }
            }


            //TODO: THIS ASSUMES YOU'LL ALWAYS STUFF LOT NUMBER, EVEN IF LOTS ALREADY SEQUENCED IN SA. PHASE II SHOULD PULL LOT NUMBERS FOR CHANGE IN CASE THEY WERE SEQUENCED
            //Stuff Lot Numbers
            string desc = "";
            //int lotnbr = 990000;
            int lotcol = 0;
            for (int c = 1; c <= colCount; c++)
            {
                if (thisWS.Cells[1, c].Value == "LotNumber")
                {
                    lotcol = c;
                    break;
                }
            }

            int desccol = 0;
            for (int c = 1; c <= colCount; c++)
            {
                if (thisWS.Cells[1, c].Value == "Description")
                {
                    desccol = c;
                    break;
                }
            }

            if (lotcol != 0 && desccol != 0)
            {
                for (int r = 2; r <= rowCount; r++)
                {
                    //thisWS.Cells[r, lotcol].Value = lotnbr;
                    //lotnbr += 1;

                    desc = "";

                    cmd2.CommandText = "select descrip from SAN_Sale_Data  where LOT_NO = '" + thisWS.Cells[r, lotcol].Value + "'";    //mapping step stuffed CM value, so now re-map
                    reader2 = cmd2.ExecuteReader();
                    
                    if (reader2.HasRows)
                    {
                        while (reader2.Read())
                        {

                            //CMCategoryTxt = reader2.GetString(1);
                            //SACategoryTxt = reader2.GetString(3);
                            //EBCategoryTxt = reader2.GetString(5);

                            desc = reader2.GetString(0);              //Note EBay id is string until further known
                            thisWS.Cells[r, desccol].Value = desc;
                            thisWS.Cells[r, desccol].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup

                        }
                    }
                    else
                    {
                        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                        thisWS.Cells[r, desccol].Interior.Color = Color.Red;

                        //TODO: add row,column and heading to comment
                        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                        thisWS.Cells[r, desccol].ClearComments();
                        thisWS.Cells[r, desccol].AddComment("Tried to find description DESCRIP in the SAN data for this lots: " + thisWS.Cells[r, lotcol].Value );
                        //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                    }

                    reader2.Close();





                }
            }





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

        private void ValidateSpreadsheetAMS()
        {
            //((Excel.Range) thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            Globals.ThisAddIn.Application.StatusBar = String.Format("All files loaded - Now we are validating the target SimpleAuction spreadsheet. Please be patient.");

            Cursor.Current = Cursors.WaitCursor;

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //*****************************************************************************************************************************************************
            //* Load up the database driven mapped fields: categor, sale, consignor, consignment
            //****************************************************************************************************************************************************


            //TODO: protect ranges where you don't want them to change directly


            //Set up connection for multiple queries
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL\\SQLEXPRESS ;Initial Catalog=SimpleAuction;Integrated Security=True");
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection2 = new SqlConnection("Data Source=MANCINI-AWARE;Initial Catalog=Describing;Integrated Security=True");
            SqlCommand cmd2 = new SqlCommand();
            cmd2.CommandType = CommandType.Text;
            cmd2.Connection = sqlConnection2;
            sqlConnection2.Open();

            SqlDataReader reader2;



            //TODO: MAKE SURE GETTING LAST ROW AND LAST COLUMN

            //****************************************************************************************************************************************************************************
            //* Validate Required fields (indicated in ExcelHeadingMap
            //****************************************************************************************************************************************************************************
            Excel.Range thisRange = thisWS.UsedRange;
            int rowCount = thisRange.Rows.Count;
            int colCount = thisRange.Columns.Count;

            //build list of SA columns that are required
            //ArrayList reqSaColNbr = new ArrayList();
            ArrayList reqSaColName = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.Required)           //if CM heading mapped to SA Heading
                {
                    //reqSaColNbr.Add(map.SAPosition);
                    reqSaColName.Add(map.SAHead);

                }
            }

            //get count of required columns to walk
            int requiredCount = reqSaColName.Count;
            //int requiredCount = reqSaColNbr.Count;

            //build list of SA columns that are mapped

            //ArrayList mapSaColNbr = new ArrayList();
            ArrayList mapSaColName = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.mapDB)           //if CM heading mapped to SA Heading
                {
                    //mapSaColNbr.Add(map.SAPosition);
                    mapSaColName.Add(map.SAHead);
                }
            }




            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //use the current row in the "TO" spreadsheet- (outer loop)
            //((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Interior.Color = Color.White;
            // ((Excel.Range)thisWS.Cells[rowCount, colCount]).ClearComments();
            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Clear();   //SHOULD RESET COLOR AND COMMENTS INSTEAD OF ABOVE


            //*********************************************************************************************************************************************************
            //* Edit for required fields first
            //*********************************************************************************************************************************************************
            //System.Type type;

            for (int r = 2; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    //value = thisWS.Cells[r, c].Value;
                    //valueint = 0;
                    //o = Int32.TryParse(value, out valueint);


                    foreach (string reqSAName in reqSaColName)
                    {
                        //value = thisWS.Cells[r, c].Value;
                        //if ((c == reqSACol) && ((thisWS.Cells[r, c].Value == null) || (valueint == 0))) {

                        //TODO --> type = (thisWS.Cells[r, c].Value).GetType;  //TODO: NEED TO CHECK FOR 0 IN SOME FIELDS I.E. SEQUENCE NBR
                        if (thisWS.Cells[1, c].Value == reqSAName && thisWS.Cells[r, c].Value == null)
                        {
                            ((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;

                            //TODO: add row,column and heading to comment
                            String txt = thisWS.Cells[1, c].Value;
                            thisWS.Cells[r, c].ClearComments();
                            thisWS.Cells[r, c].AddComment(txt + " is required");
                            //((Excel.Range)ws.Cells[r, c]).Style.Name = "Normal"
                        }
                    }
                }
            }

            //****************************************************************************************************************************************************************************
            // Field Mapping exercises - IS Field title specific - no longer relies on column number
            //****************************************************************************************************************************************************************************

            int SACategoryId = 0;
            string EbayCategoryId = "";
            int SAAuctionId = 0;
            int SAConsignorId = 0;
            int SAConsignmentId = 0;
            int SAPriceGuide1Id = 0;
            int SAPriceGuide2Id = 0;
            int SALOAProviderId = 0;
            int SAShippingCategoryId = 0;

            //Check fields that require database mapping by ID
            for (int r = 2; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    if ((thisWS.Cells[1, c].Value != null) && (thisWS.Cells[r, c].Value != null))   //mapped columns are required so will be red if not provided
                    {
                        //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                        //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                        //((Excel.Range)thisWS.Cells[r, reqSaColNbr[c]]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                        if (thisWS.Cells[1, c].Value == "AuctionID")
                        {
                            SAAuctionId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Auction where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAAuctionId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAAuctionId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM sale id: " + thisWS.Cells[r, c].Value + " to SA auction id - CM sale id not found in Auction table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "CategoryId")
                        {
                            SACategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    string val = thisWS.Cells[r, c].Value;

                                    SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SACategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Country Lookup: " + val + " to SA Category SAID " + SACategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Country Lookup: " + thisWS.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "CategoryId2")
                        {
                            SACategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    string val = thisWS.Cells[r, c].Value;

                                    SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SACategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Province Lookup: " + val + " to SA Category SAID " + SACategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Province Lookup: " + thisWS.Cells[r, c].Value + " to SA category id - CM CMCategoryTxt not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "ConsignerId")
                        {
                            SAConsignorId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAConsignorId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS.Cells[r, c].Value + " to SA consignor id - CM consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        //****************
                        else if (thisWS.Cells[1, c].Value == "ConsignmentTrackingId")
                        {
                            SAConsignmentId = 0;


                            cmd2.CommandText = "SELECT SAId FROM dbo.Consignment where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAConsignmentId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAConsignmentId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM consignment id: " + thisWS.Cells[r, c].Value + " to SA consignment id - CM consignment id not found in Consignment table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "ShippingCategoryId")
                        {
                            SAShippingCategoryId = 0;

                            cmd2.CommandText = "SELECT SAid FROM dbo.Shipping_Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {
                                    string val = thisWS.Cells[r, c].Value;

                                    SAShippingCategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAShippingCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;

                                    thisWS.Cells[r, c].ClearComments();
                                    thisWS.Cells[r, c].AddComment("Mapped CM Package Type: " + val + " to SA Shipping_Category (SAID) " + SAShippingCategoryId);

                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Package Type: " + thisWS.Cells[r, c].Value + " to SA Shipping_Category id - CM CMCategoryTxt not found in Shipping_Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }

                        else if (thisWS.Cells[1, c].Value == "PriceGuide1Id")
                        {
                            SAPriceGuide1Id = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAPriceGuide1Id = reader2.GetInt32(0);         //assume it's not red already because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAPriceGuide1Id;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM PriceGuide1 id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "PriceGuide2Id")
                        {
                            SAPriceGuide2Id = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.Catalog_reference where SA_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SAPriceGuide2Id = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SAPriceGuide2Id;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM PriceGuide2 id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "LOAProviderId")
                        {
                            SALOAProviderId = 0;

                            cmd2.CommandText = "SELECT SAId FROM dbo.LOA_Provider where CM_Name = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    SALOAProviderId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                    thisWS.Cells[r, c].Value = SALOAProviderId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                    //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM Certificate Id: " + thisWS.Cells[r, c].Value + " to SA price guide id - CM PriceGuide id not found in Catalog_Reference table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "EbayPrimaryCategoryId" && EBayImplemented)
                        {
                            EbayCategoryId = "";

                            cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                                    thisWS.Cells[r, c].Value = EbayCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                        //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }

                        // https://drive.google.com/drive/folders/1grl8P1eV5HUsjd0_LlLPVHJfh9eukaPz
                        else if (thisWS.Cells[1, c].Value == "Condition")
                        {
                            string symbols = thisWS.Cells[r, c].Value;
                            var imgs = new List<string>();


                            int i = 0;
                            while ((i = symbols.IndexOf("img", i)) != -1)
                            {
                                // Print out the substring.
                                //Console.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i,3), i);
                                Debug.WriteLine("row: {0} has {1} in position {2}", r, symbols.Substring(i, 3), i);

                                //<img src="http://www.kelleherauctions.com/images/mint.gif" align="top">
                                string imgname = "";
                                int start = i + 45;
                                int backslash = symbols.IndexOf("/", start);
                                backslash += 1;
                                int period = symbols.IndexOf(".", start);
                                int displace = 0;

                                if (period != -1)
                                {
                                    displace = period - backslash;

                                    imgname = symbols.Substring(backslash, displace);
                                    Debug.WriteLine("row: {0} image name {1}", r, imgname);

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
                                    if (!inbracket && ch != '/')
                                {
                                    term += ch;
                                }


                            }

                            if (term != "")
                            {
                                imgs.Add(term);
                                term = "";
                            }

                            string imgid = "";
                            string origimg = "";
                            if (imgs.Count > 0)
                            {

                                origimg = thisWS.Cells[r, c].Value;
                                thisWS.Cells[r, c].Value = ""; //clear out value
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Original condition before xform: " + origimg);


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

                                    cmd2.CommandText = "select id from symbol_reference where cmid = '" + s + "'";    //mapping step stuffed CM value, so now re-map
                                    reader2 = cmd2.ExecuteReader();


                                    if (reader2.HasRows)
                                    {
                                        while (reader2.Read())
                                        {

                                            //CMCategoryTxt = reader2.GetString(1);
                                            //SACategoryTxt = reader2.GetString(3);
                                            //EBCategoryTxt = reader2.GetString(5);

                                            imgid = reader2.GetString(0);              //Note EBay id is string until further known
                                            thisWS.Cells[r, c].Value = thisWS.Cells[r, c].Value + imgid + ";";
                                            thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                                //break;
                                        }
                                    }
                                    else
                                    {
                                        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                        thisWS.Cells[r, c].Interior.Color = Color.Red;

                                        //TODO: add row,column and heading to comment
                                        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                        thisWS.Cells[r, c].ClearComments();
                                        thisWS.Cells[r, c].AddComment("Tried to find image for CMid " + s + " in table Symbol_Reference. CMid not found");
                                        //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                    }

                                    reader2.Close();
                                }
                            }

                            //string sub = input.Substring(0, 3);
                            //Console.WriteLine("Substring: {0}", sub);

                        }

                        else if (thisWS.Cells[1, c].Value == "EbaySecondaryCategoryId" && EBayImplemented)
                        {
                            EbayCategoryId = "";

                            cmd2.CommandText = "SELECT EBid FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                            reader2 = cmd2.ExecuteReader();


                            if (reader2.HasRows)
                            {
                                while (reader2.Read())
                                {

                                    //CMCategoryTxt = reader2.GetString(1);
                                    //SACategoryTxt = reader2.GetString(3);
                                    //EBCategoryTxt = reader2.GetString(5);

                                    EbayCategoryId = reader2.GetString(0);              //Note EBay id is string until further known
                                    thisWS.Cells[r, c].Value = EbayCategoryId;
                                    thisWS.Cells[r, c].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup
                                                                                        //break;
                                }
                            }
                            else
                            {
                                //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //TODO: add row,column and heading to comment
                                //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                thisWS.Cells[r, c].ClearComments();
                                thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to EBay category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                            }

                            reader2.Close();
                        }
                        else if (thisWS.Cells[1, c].Value == "SerialNumber")
                        {

                            if (thisWS.Cells[r, c].Value != null && thisWS.Cells[r, c].Value != 0)
                            {

                                //var client = new RestClient("http://kelleherdemo2-com.si-sv2521.com");
                                var client = new RestClient("http://kelleher-stage-com.si-sv2521.com/Kelleher.aspx?debug=GetInventoryIdBySerialNumber");

                                var request = new RestRequest("/Kelleher.aspx", Method.POST);
                                request.RequestFormat = DataFormat.Json;

                                double serialNbrStr = thisWS.Cells[r, c].Value;
                                //string serialNbrStr = thisWS.Cells[r, c].Value;
                                //string serialNbrStr = serialNbr.ToString();

                                //https://stackoverflow.com/questions/14828520/how-to-create-my-json-string-by-using-c
                                //var f = new SARestLoginModel
                                //{
                                //    request = new Dictionary<string, string>
                                //    {
                                //        {"username", "admin"},
                                //        {"password", "admin"},
                                //        {"operation", "GetConsignors"},
                                //        //{"serialnumber", serialNbrStr},
                                //    }
                                //};

                                //{"request":{"username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}}
                                //request.AddJsonBody(new { A = "foo", B = "bar" });
                                //request.AddJsonBody(new { "request":{ "username":"admin","password":"admin","operation":"GetInventoryIdBySerialNumber", "serialnumber":"serialnumber"}
                                request.AddJsonBody(new { request = new { username = "admin", password = "admin", operation = "GetInventoryIdBySerialNumber", serialnumber = serialNbrStr.ToString() } });

                                //});

                                //request.AddBody(f);
                                //request.AddXmlBody(f);

                                IRestResponse response = client.Execute(request);

                                if (!response.IsSuccessful)
                                {
                                    Debug.WriteLine("reponse failed");
                                }

                                JObject obj1 = JObject.Parse(response.Content);
                                //JArray SAInventoryId = (JArray)obj1["inventoryid"];
                                JValue SAInventoryId = (JValue)obj1["inventoryid"];

                                //int len = SAInventoryId.Count;
                                //int inventoryId = 0;
                                string inventoryId = (string)SAInventoryId;
                                //inventoryId = (int)SAInventoryId[0]["id"]; //????


                                if (inventoryId != null && inventoryId != "-1")            //TODO: CHECK FOR -1
                                {
                                    //int col = GetInvColumn(thisWS,r);

                                    for (int i = 1; i < 256; i++)
                                    {
                                        if (thisWS.Cells[1, i].Value == "InventoryId")
                                        {
                                            //return (i);
                                            thisWS.Cells[r, i].Value = inventoryId; //TODO BREAK WHEN HIT FIRST ONE
                                            break;
                                        }
                                    }
                                }


                                //thisWS.Cells[r, c].Value = EbayCategoryId;
                                //thisWS.Cells[r, c].Interior.Color = Color.Blue;


                                //SAInventoryId = 0;

                                //cmd2.CommandText = "SELECT SAId FROM dbo.Consignor where CMId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                                //reader2 = cmd2.ExecuteReader();


                                //if (reader2.HasRows)
                                //{
                                //    while (reader2.Read())
                                //    {

                                //        //CMCategoryTxt = reader2.GetString(1);
                                //        //SACategoryTxt = reader2.GetString(3);
                                //        //EBCategoryTxt = reader2.GetString(5);

                                //        SAConsignorId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                //        thisWS.Cells[r, c].Value = SAConsignorId;
                                //        thisWS.Cells[r, c].Interior.Color = Color.Blue;
                                //        //break;
                                //    }
                                //}
                                //else
                                //{
                                //    //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                                //    thisWS.Cells[r, c].Interior.Color = Color.Red;

                                //    //TODO: add row,column and heading to comment
                                //    //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                                //    thisWS.Cells[r, c].ClearComments();
                                //    thisWS.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS.Cells[r, c].Value + " to SA consignor id - CM consignor id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                //    //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                //}

                                //reader2.Close();
                            }
                        }






                        //**** Map Consignor Id ****
                        //if (thisWS.Cells[1, c].Value == "CategoryId")
                        //{
                        //    SACategoryId = 0;

                        //    cmd2.CommandText = "SELECT * FROM dbo.Category where CMCategoryTxt = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
                        //    reader2 = cmd2.ExecuteReader();


                        //    if (reader2.HasRows)
                        //    {
                        //        while (reader2.Read())
                        //        {

                        //            //CMCategoryTxt = reader2.GetString(1);
                        //            //SACategoryTxt = reader2.GetString(3);
                        //            //EBCategoryTxt = reader2.GetString(5);

                        //            SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                        //            thisWS.Cells[r, c].Value = SACategoryId;
                        //            thisWS.Cells[r, c].Interior.Color = Color.Blue;
                        //            break;
                        //        }
                        //    }
                        //    else
                        //    {
                        //        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                        //        thisWS.Cells[r, c].Interior.Color = Color.Red;

                        //        //TODO: add row,column and heading to comment
                        //        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                        //        thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to SA category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                        //    }



                        //    reader2.Close();
                        //}



                    }
                }
            }


            //TODO: THIS ASSUMES YOU'LL ALWAYS STUFF LOT NUMBER, EVEN IF LOTS ALREADY SEQUENCED IN SA. PHASE II SHOULD PULL LOT NUMBERS FOR CHANGE IN CASE THEY WERE SEQUENCED
            //Stuff Lot Numbers
            string desc = "";
            //int lotnbr = 990000;
            int lotcol = 0;
            for (int c = 1; c <= colCount; c++)
            {
                if (thisWS.Cells[1, c].Value == "LotNumber")
                {
                    lotcol = c;
                    break;
                }
            }

            int desccol = 0;
            for (int c = 1; c <= colCount; c++)
            {
                if (thisWS.Cells[1, c].Value == "Description")
                {
                    desccol = c;
                    break;
                }
            }

            if (lotcol != 0 && desccol != 0)
            {
                for (int r = 2; r <= rowCount; r++)
                {
                    //thisWS.Cells[r, lotcol].Value = lotnbr;
                    //lotnbr += 1;

                    desc = "";

                    cmd2.CommandText = "select descrip from SAN_Sale_Data  where LOT_NO = '" + thisWS.Cells[r, lotcol].Value + "'";    //mapping step stuffed CM value, so now re-map
                    reader2 = cmd2.ExecuteReader();

                    if (reader2.HasRows)
                    {
                        while (reader2.Read())
                        {

                            //CMCategoryTxt = reader2.GetString(1);
                            //SACategoryTxt = reader2.GetString(3);
                            //EBCategoryTxt = reader2.GetString(5);

                            desc = reader2.GetString(0);              //Note EBay id is string until further known
                            thisWS.Cells[r, desccol].Value = desc;
                            thisWS.Cells[r, desccol].Interior.Color = Color.Blue;     //assume it's not red alread because these was a value to lookup

                        }
                    }
                    else
                    {
                        //((Excel.Range)thisWS.Cells[r, c]).Interior.Color = Color.Red;
                        thisWS.Cells[r, desccol].Interior.Color = Color.Red;

                        //TODO: add row,column and heading to comment
                        //((Excel.Range)thisWS.Cells[r, x]).AddComment(thisWS.Cells[r, c].Value + " is required") ;
                        thisWS.Cells[r, desccol].ClearComments();
                        thisWS.Cells[r, desccol].AddComment("Tried to find description DESCRIP in the SAN data for this lots: " + thisWS.Cells[r, lotcol].Value);
                        //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                    }

                    reader2.Close();





                }
            }





            Cursor.Current = Cursors.Default;

        }


        private void btnVerify_Click_1(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You hit the verify button");



        }
    }
}

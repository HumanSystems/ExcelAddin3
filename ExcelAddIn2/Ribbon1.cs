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

        }


        //******************************************************************************************************************************************
        //Get reference to current sheet (with add-in) that will hold resulting SA Lot file
        //*****************************************************************************************************************************************


        ArrayList headingsMap = new ArrayList();

        

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnLoadCatMast_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS.UsedRange;
            int thisRowCount = thisRange.Rows.Count;
            int thisColCount = thisRange.Columns.Count;

            // thisWS.Cells[thisRowCount, thisColCount].Clear();               //this bullshit doesn't work
            // thisWS.Cells[thisRowCount, thisColCount].ClearComments();        //this bullshit doesn't work

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
            openFileDialog1.Filter = "xls Files|*.xls|xlsx Files|*.xlsx";
            openFileDialog1.Title = "Select Source Catalog Master Excel Lot File for Simple Auction";

            //Open file selection dialog - if canceled out just return. Otherwise perform all processing to suck in the selected Catalog Master 
            //lot export file
            if (openFileDialog1.ShowDialog() != System.Windows.Forms.DialogResult.OK)
            {
                return;
            }


            //TODO: CHECK FILE NAME BEFORE OPENING AGAINST ??? TO MAKE SURE IT'S AN ?UNPROCESSED? ?NEW? ?WELL-NAMED? CATALOG MASTER FILE
            string filename = openFileDialog1.FileName;
            MessageBox.Show("here is selected file name: " + filename);

           

            //1.A.1 - SQL to retrieve Table ExcelHeadingMap ordered by SA Column Position (starting with 1)
           
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
            SqlConnection sqlConnection1 = new SqlConnection("Data Source=BACKUPDELL ;Initial Catalog=Describing;Integrated Security=True");
            SqlCommand cmd1 = new SqlCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.Connection = sqlConnection1;
            SqlDataReader reader1;
            cmd1.CommandText = "SELECT SAColumnNbr,SAHeading,CMColumnNbr,CMHeading, Required, DefaultValue, mapDB FROM dbo.ExcelHeadingMap order by SAColumnNbr";

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

                    //Populate first row current column with SA heading
                    ((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;

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

            reader1.Close();
            cmd1.Dispose();
            sqlConnection1.Close();
            sqlConnection1.Dispose();

            //TODO: CHECK DATA BY COLUMN AS IMPORT

            //****************************************************************************************************************************************************
            //* after 1) IMPORT AND 2) GRAB SA/EBAY CATEGORY IDS FROM MAPPING AND 3) APPLY BUSINESS RULES FOR SUB-CATEGORY SORT FIELDS 4) GRAB CONSIGNOR, SALE#  
            //TODO: CRITICAL - GRA



            var fromXlApp = new Excel.Application();
            //xlApp.Visible = true;
            fromXlApp.Visible = false; //--> Don't need to see the Catalog Master excel file to suck it in
                
            Excel.Workbook fromXlWorkbook = fromXlApp.Workbooks.Open(filename);     //this is the fully qualified (local) file name
            Excel._Worksheet fromXlWorksheet = fromXlWorkbook.Sheets[1];            //TODO: make sure only one worksheet???

            Excel.Range fromXlRange = fromXlWorksheet.UsedRange;
                

            MessageBox.Show("CM (from) file should be open now ... begin data map/load from CM to current SA lot spreadsheet");

            Cursor.Current = Cursors.WaitCursor;

            //TODO: MAKE SURE COUNTS ARE NOT ENTIRE WORKSHEET
            int rowCount = fromXlRange.Rows.Count;
            int colCount = fromXlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            //1) Walk table ExcelHeadingMap 

            //string msg = "";

           

            //WALK THE TO SPREADSHEET ROWS AND POPULATE WITH FROM SPREADSHEET VALUES BASED ON MAPPING (ExcelHeadingMap)
            for (int r = 2; r <= rowCount; r++)         //r = TO ROW TO FILL - WALK HEADING MAP ARRAY TO LEARN COLUMNS TO COPY
            {
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
                        ((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE spreadsheet to spreadsheet VALUE GET'S MOVED!!!

                        }
                        else
                            if (map.SAHead != "" && map.CMHead == "" && map.defaultValue != "")  //otherwise, if there is a default value stuff it into the sa column
                            {
                                //NOTE: Default value will trump mapping
                                ((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = map.defaultValue; //THIS IS WHERE load default value from ExcelHeadingMap!!!
                            }
                }
            }

            

            //*************************************************************************************************************************************************************
            //* This will check for missing values and also map database (id) values for category, auction/sale, consignor, consignment
            //*************************************************************************************************************************************************************
            ValidateSpreadsheet();




        //TODO: cleanup EXCEL and connections
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(fromXlRange);
            Marshal.ReleaseComObject(fromXlWorksheet);

            //close and release
            fromXlWorkbook.Close();
            Marshal.ReleaseComObject(fromXlWorkbook);

        //quit and release
            fromXlApp.Quit();
            Marshal.ReleaseComObject(fromXlApp);


            Cursor.Current = Cursors.Default;


        }


        //git change again
        
        private void ValidateSpreadsheet()
        {
            //((Excel.Range) thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            Cursor.Current = Cursors.WaitCursor;

            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            
            //*****************************************************************************************************************************************************
            //* Load up the database driven mapped fields: categor, sale, consignor, consignment
            //****************************************************************************************************************************************************


            //TODO: protect ranges where you don't want them to change directly


            //Set up connection for multiple queries
            //SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL\\SQLEXPRESS ;Initial Catalog=SimpleAuction;Integrated Security=True");
            SqlConnection sqlConnection2 = new SqlConnection("Data Source=BACKUPDELL;Initial Catalog=Describing;Integrated Security=True");
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
            ArrayList reqSaColNbr = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.Required)           //if CM heading mapped to SA Heading
                {
                    reqSaColNbr.Add(map.SAPosition);
                }
            }

            //get count of required columns to walk
            int requiredCount = reqSaColNbr.Count;

            //build list of SA columns that are mapped

            ArrayList mapSaColNbr = new ArrayList();
            foreach (OneColumnMap map in headingsMap)
            {
                if (map.mapDB)           //if CM heading mapped to SA Heading
                {
                    mapSaColNbr.Add(map.SAPosition);
                }
            }




            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //((Excel.Range)thisWS.Cells[1, thisColumnMap.SAPosition]).Value = thisColumnMap.SAHead;
            //use the current row in the "TO" spreadsheet- (outer loop)
            //((Excel.Range)thisWS.Cells[r, map.SAPosition]).Value = (fromXlRange.Cells[r, map.CMPosition].Value); //THIS IS WHERE VALUE GET'S MOVED!!!

            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Interior.Color = Color.White;
            // ((Excel.Range)thisWS.Cells[rowCount, colCount]).ClearComments();
            //((Excel.Range)thisWS.Cells[rowCount, colCount]).Clear();   //SHOULD RESET COLOR AND COMMENTS INSTEAD OF ABOVE


            int x = 0;

            for (int r = 2; r <= rowCount; r++) {
                for (int c = 1; c <= colCount; c++) {
                    foreach (int reqSACol in reqSaColNbr) {
                        if ((c == reqSACol) && ((thisWS.Cells[r, c].Value == null))) {
                            //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                            //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                            //((Excel.Range)thisWS.Cells[r, reqSaColNbr[c]]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                            x = reqSACol;  //TODO: IS THIS NEEDED?
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

            int SACategoryId = 0;
            string EbayCategoryId = "";
            int SAAuctionId = 0;
            int SAConsignorId = 0;


            //Walk spreadsheet again to get mapped fields from database
            for (int r = 2; r <= rowCount; r++) {
                for (int c = 1; c <= colCount; c++) {
                    foreach (int mapSACol in mapSaColNbr) { 
                        if ((c == mapSACol) && (thisWS.Cells[r, c].Value != null))   //mapped columns are required so will be red if not provided
                        {
                            //((Excel.Range)ws.Cells[r, c]).NumberFormat = format;
                            //((Excel.Range)ws.Cells[r, c]).Value2 = cellVal;
                            //((Excel.Range)thisWS.Cells[r, reqSaColNbr[c]]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                            if (thisWS.Cells[1, c].Value == "AuctionID")
                            {
                                SAAuctionId = 0;

                                cmd2.CommandText = "SELECT SAAuctionId FROM dbo.Auction where CMAuctionId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
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

                                        SACategoryId = reader2.GetInt32(0);         //assume it's not red alread because these was a value to lookup
                                        thisWS.Cells[r, c].Value = SACategoryId;
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
                                    thisWS.Cells[r, c].AddComment("Tried to map CM category: " + thisWS.Cells[r, c].Value + " to SA category id - CM category not found in Category table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                    //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                }

                                reader2.Close();
                            }
                            else if (thisWS.Cells[1, c].Value == "ConsignerId")
                            {
                                SAConsignorId = 0;

                                cmd2.CommandText = "SELECT SAConsignorId FROM dbo.Consignor where CMConsignorId = '" + thisWS.Cells[r, c].Value + "'";    //mapping step stuffed CM value, so now re-map
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
                                    thisWS.Cells[r, c].AddComment("Tried to map CM consignor id: " + thisWS.Cells[r, c].Value + " to SA consignor id - CM sale id not found in Consignor table. Mapping is required - please add mapping to table and re-validate this spreadsheet");
                                    //thisWS.Cells[r, c].Comment[1].AutoFit = true;
                                }

                                reader2.Close();
                            }
                            else if (thisWS.Cells[1, c].Value == "EbayPrimaryCategoryId")
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
                            else if (thisWS.Cells[1, c].Value == "EbaySecondaryCategoryId")
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
            }

            Cursor.Current = Cursors.Default;

        }


            private void btnVerify_Click_1(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("You hit the verify button");



        }
    }
}

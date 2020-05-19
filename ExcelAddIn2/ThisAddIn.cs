using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;


namespace ExcelAddIn2
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
        }

        void Application_WorkbookBeforeSave(Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
        {
            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range thisRange = thisWS.UsedRange;
            int rowCount = thisRange.Rows.Count;
            //int colCount = thisRange.Columns.Count;


            if (rowCount < 2)
            {
                Cancel = true;
                MessageBox.Show("You haven't loaded data yet - please load data before you save anythin");
                return;
            }


            if (Ribbon1.nbrFatalErrors != 0)
            {
                MessageBox.Show("WARNING! Workbook has " + Ribbon1.nbrFatalErrors.ToString() + " errors - please do not import it into AMS");
            }
            else
            {
                MessageBox.Show("Workbook has 0 errors and is ready to import into AMS");
            }


            //if (DialogResult.No == MessageBox.Show("Are you sure you want to " +
            //    "save the workbook?", "Example", MessageBoxButtons.YesNo))
            //{
            //    Cancel = true;
            //    MessageBox.Show("Save is canceled.");
            //}
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

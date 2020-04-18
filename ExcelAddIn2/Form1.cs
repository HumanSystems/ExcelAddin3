using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {





            Excel.Worksheet thisWS = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (thisWS.AutoFilterMode )
            {
                MessageBox.Show("Autofilter is on ");
            }
            else
            {
                MessageBox.Show("Autofilter is off ");
                return;
            }



            Excel.Range visibleCells = thisWS.UsedRange.SpecialCells(
                               Excel.XlCellType.xlCellTypeVisible,
                               Type.Missing);

            foreach (Excel.Range area in visibleCells.Areas)
            {
                foreach (Excel.Range row in area.Rows)
                {
                    if (row.Cells[1, 2].Value2 != null)
                    {
                        MessageBox.Show(String.Format("The row value for row number {0} ",
                         Convert.ToString(row.Cells[1, 2].Value2)));
                    }
                    else
                    {
                        break;
                    }
                }
            }

            TreeNode treeNode = new TreeNode("Stamps");
            treeView1.Nodes.Add(treeNode);

            treeNode = new TreeNode("Coins");
            treeView1.Nodes.Add(treeNode);

        }
    }
}

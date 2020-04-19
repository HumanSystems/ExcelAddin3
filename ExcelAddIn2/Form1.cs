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
using System.Data.SqlClient;


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

            //*********************** BEGIN TEST STUFF ***********************************************
            //TreeNode treeNode = new TreeNode("Stamps");
            //treeNode.Tag = "Cat001";
            //treeView1.Nodes.Add(treeNode);

            //treeNode = new TreeNode("Coins");
            //treeView1.Nodes.Add(treeNode);

            //TreeNode node2 = new TreeNode("C#");
            //TreeNode node3 = new TreeNode("VB.NET");
            //TreeNode[] array = new TreeNode[] { node2, node3 };
            ////
            //// Final node.
            ////
            //treeNode = new TreeNode("Dot Net Perls", array);
            //treeView1.Nodes.Add(treeNode);

            //*********************** BEGIN REAL STUFF ***********************************************
            //https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.treeview.checkboxes?view=netframework-4.5.2
            //TODO: don't make checkboxes until determine if can have 1-X category codes
            //treeView1.CheckBoxes = true;

            //TODO: decide if whnt to show expanded
            //treeView1.ExpandAll()
            //treeView1.CollapseAll();


            SqlConnection sqlConnection1 = new SqlConnection("Data Source=MANCINI-AWARE ;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection2 = new SqlConnection("Data Source=MANCINI-AWARE ;Initial Catalog=Describing;Integrated Security=True");
            SqlConnection sqlConnection3 = new SqlConnection("Data Source=MANCINI-AWARE ;Initial Catalog=Describing;Integrated Security=True");

            SqlCommand cmd1 = new SqlCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.Connection = sqlConnection1;
            SqlDataReader reader1;
            cmd1.CommandText = "SELECT id,  Sequence, AMSid, Amstxt FROM dbo.Category_AMS where ParentId is null and Active = 1 order by Sequence";
            sqlConnection1.Open();
            reader1 = cmd1.ExecuteReader();

            if (reader1.HasRows)
            {
                
                while (reader1.Read())
                {
                                       
                    int id1 = reader1.GetInt32(0);
                    int sequence1 = reader1.GetInt32(1);
                    string amsid1 = (reader1.IsDBNull(2) ? null : reader1.GetString(2));
                    //string amsid1 = (reader1.IsDBNull(2) ? "" : reader1.GetString(2));
                    string amstxt1 = reader1.GetString(3);

                    TreeNode treeNode = new TreeNode(); 
                    if (amsid1 != null)
                    {
                        treeNode.Tag = amsid1;
                        treeNode.Text = "[" + amsid1 + "] " + amstxt1;;
                    }
                    else
                    {
                        treeNode.Text = amstxt1; 
                    }

                    int n = treeView1.Nodes.Add(treeNode);  //TODO: the int here is the index added - don't need if don't use

                    //--> this adds a child node to the current node: treeNode.Nodes.Add("childDevice");
                    SqlCommand cmd2 = new SqlCommand();
                    cmd2.CommandType = CommandType.Text;
                    cmd2.Connection = sqlConnection2;   //need a second connection for some stupid reason
                    SqlDataReader reader2;
                    cmd2.CommandText = "SELECT id,  Sequence, AMSid, Amstxt FROM dbo.Category_AMS where ParentId = " + id1.ToString() + " and Active = 1 order by Sequence";
                    sqlConnection2.Open();
                    reader2 = cmd2.ExecuteReader();

                    while (reader2.Read())
                    {
                        int id2 = reader2.GetInt32(0);
                        int sequence2 = reader2.GetInt32(1);
                        //string amsid2 = (reader2.IsDBNull(2) ? "" : reader2.GetString(2));
                        string amsid2 = (reader2.IsDBNull(2) ? null : reader2.GetString(2));
                        string amstxt2 = reader2.GetString(3);

                        TreeNode treeNode2 = new TreeNode();
                        if (amsid2 != null)
                        {
                            treeNode2.Tag = amsid2;
                            treeNode2.Text = "[" + amsid2 + "] " + amstxt2; ;
                        }
                        else
                        {
                            treeNode2.Text = amstxt2;
                        }



                        //--> this adds a child node to the current node: treeNode.Nodes.Add("childDevice");
                        SqlCommand cmd3 = new SqlCommand();
                        cmd3.CommandType = CommandType.Text;
                        cmd3.Connection = sqlConnection3;   //need a second connection for some stupid reason
                        SqlDataReader reader3;
                        cmd3.CommandText = "SELECT id,  Sequence, AMSid, Amstxt FROM dbo.Category_AMS where ParentId = " + id2.ToString() + " and Active = 1 order by Sequence";
                        sqlConnection3.Open();
                        reader3 = cmd3.ExecuteReader();

                        while (reader3.Read())
                        {
                            int id3 = reader3.GetInt32(0);
                            int sequence3 = reader3.GetInt32(1);
                            //string amsid3 = (reader3.IsDBNull(2) ? "" : reader3.GetString(2));
                            string amsid3 = (reader3.IsDBNull(2) ? null : reader3.GetString(2));
                            string amstxt3 = reader3.GetString(3);

                            TreeNode treeNode3 = new TreeNode();
                            if (amsid3 != null)
                            {
                                treeNode3.Tag = amsid3;
                                treeNode3.Text = "[" + amsid3 + "] " + amstxt3; ;
                            }
                            else
                            {
                                treeNode3.Text = amstxt3;
                            }

                            treeNode2.Nodes.Add(treeNode3);   //need to add third level node to send level node before adding second level to current (root) node 
                        }

                        treeNode.Nodes.Add(treeNode2);


                        sqlConnection3.Close();
                    }
                    sqlConnection2.Close();


                    ////thisColumnMap.CMPosition = reader1.GetInt32(2);
                    //thisColumnMap.CMPosition = (reader1.IsDBNull(2) ? 0 : reader1.GetInt32(2));
                    ////thisColumnMap.CMHead = reader1.GetString(3);
                    //thisColumnMap.CMHead = (reader1.IsDBNull(3) ? "" : reader1.GetString(3));
                    //thisColumnMap.Required = (reader1.IsDBNull(4) ? false : reader1.GetBoolean(4));
                    ////thisColumnMap.Required = reader1.GetBoolean(4);
                    //thisColumnMap.defaultValue = (reader1.IsDBNull(5) ? "" : reader1.GetString(5));
                    //thisColumnMap.mapDB = (reader1.IsDBNull(6) ? false : reader1.GetBoolean(6));
                    //thisColumnMap.SARequired = (reader1.IsDBNull(7) ? false : reader1.GetBoolean(7));
                    //thisColumnMap.Note = (reader1.IsDBNull(8) ? "" : reader1.GetString(8));
                }
                sqlConnection1.Close();
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
                // Get the selected node.
                //
                TreeNode node = treeView1.SelectedNode;
                //
                // Render message box.
                //
                //TODO: this is how you get a key for the node
                MessageBox.Show(string.Format("You selected: {0} with tag: {1}", node.Text, node.Tag));
        }
    }
}

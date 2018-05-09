using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using app = Microsoft.Office.Interop.Excel.Application;
namespace First
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
       //Mo file excel
        private void btnPF_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if(openFileDialog1.ShowDialog()==System.Windows.Forms.DialogResult.OK)
            {
                this.txtPF.Text = openFileDialog1.FileName;
            }

        }
        private DataSet ds;
        private DataTable dt;
        private void btnLoadFile_Click(object sender, EventArgs e)
        {
         
            try
            {
                OleDbConnection connection = new OleDbConnection();
                connection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + txtPF.Text); //Excel 97-2003, .xls
                string excelQuery = @"Select * FROM [" + txtSN.Text + "$]";
                connection.Open();
                OleDbCommand cmd = new OleDbCommand(excelQuery, connection);
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                adapter.SelectCommand = cmd;
                ds = new DataSet();
                adapter.Fill(ds);
                dt = ds.Tables[0];
                dataGV1.DataSource = dt.DefaultView;
                //them hang
                foreach (DataRow dr in dt.Rows )
                {
                    int n = dataGV2.Rows.Add();
                    dataGV2.Rows[n].Cells[0].Value = dr[0].ToString();
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dataGV2.Rows[i].Cells[1].Value = 99;
                    dataGV2.Rows[i].Cells[2].Value = dataGV1.Rows[i].Cells[1].Value.ToString();
                    dataGV2.Rows[i].Cells[3].Value = "N";
                    dataGV2.Rows[i].Cells[4].Value = "F";
                    dataGV2.Rows[i].Cells[5].Value = "F";
                    dataGV2.Rows[i].Cells[6].Value = "";
                    dataGV2.Rows[i].Cells[7].Value = "A";
                    dataGV2.Rows[i].Cells[8].Value = "";
                    dataGV2.Rows[i].Cells[9].Value = "";
                    dataGV2.Rows[i].Cells[10].Value = "";
                    dataGV2.Rows[i].Cells[11].Value = "";
                    dataGV2.Rows[i].Cells[12].Value = "";
                    dataGV2.Rows[i].Cells[13].Value = "";
                    dataGV2.Rows[i].Cells[14].Value = "";
                    dataGV2.Rows[i].Cells[15].Value = "";
                    dataGV2.Rows[i].Cells[16].Value = "";
                    dataGV2.Rows[i].Cells[17].Value = dataGV1.Rows[i].Cells[10].Value.ToString();
                    dataGV2.Rows[i].Cells[18].Value = "";
                    dataGV2.Rows[i].Cells[19].Value = group1(dataGV1.Rows[i].Cells[37].Value.ToString());
                    dataGV2.Rows[i].Cells[20].Value = group2(dataGV1.Rows[i].Cells[37].Value.ToString());
                    dataGV2.Rows[i].Cells[21].Value = 0;
                    dataGV2.Rows[i].Cells[22].Value = 0;
                    dataGV2.Rows[i].Cells[23].Value = 1;
                    dataGV2.Rows[i].Cells[24].Value = 0;
                    dataGV2.Rows[i].Cells[25].Value = 0;
                    dataGV2.Rows[i].Cells[26].Value = 1;
                    dataGV2.Rows[i].Cells[27].Value = 0;
                    dataGV2.Rows[i].Cells[28].Value = 1;
                    dataGV2.Rows[i].Cells[29].Value = "";
                    dataGV2.Rows[i].Cells[30].Value = "";
                    dataGV2.Rows[i].Cells[31].Value = "";
                    dataGV2.Rows[i].Cells[32].Value = "";
                    dataGV2.Rows[i].Cells[33].Value = "";
                    dataGV2.Rows[i].Cells[34].Value = "";
                    dataGV2.Rows[i].Cells[35].Value = "";
                    dataGV2.Rows[i].Cells[36].Value = "";
                    dataGV2.Rows[i].Cells[37].Value = "";
                    dataGV2.Rows[i].Cells[38].Value = "";
                    dataGV2.Rows[i].Cells[39].Value = "";
                    dataGV2.Rows[i].Cells[40].Value = "UPC";
                    dataGV2.Rows[i].Cells[41].Value = dataGV1.Rows[i].Cells[22].Value.ToString();
                    dataGV2.Rows[i].Cells[42].Value = "";
                    dataGV2.Rows[i].Cells[43].Value = "";
                    dataGV2.Rows[i].Cells[44].Value = "";

                    if (dataGV1.Rows[i].Cells[7].Value.ToString() == "")
                        dataGV1.Rows[i].Cells[7].Value = 0;
                    dataGV2.Rows[i].Cells[45].Value = Int32.Parse(dataGV1.Rows[i].Cells[7].Value.ToString());

                    if (int.Parse(dataGV2.Rows[i].Cells[45].Value.ToString()) > 0)
                        dataGV2.Rows[i].Cells[46].Value = 2;
                    else
                        dataGV2.Rows[i].Cells[46].Value = null;

                    if (dataGV1.Rows[i].Cells[6].Value.ToString() == "")
                        dataGV1.Rows[i].Cells[6].Value = 0;

                    if (int.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 2)
                    {
                        dataGV2.Rows[i].Cells[47].Value = int.Parse(dataGV1.Rows[i].Cells[5].Value.ToString()) * 30;
                    }
                    else if (int.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 1)
                    {
                        dataGV2.Rows[i].Cells[47].Value = int.Parse(dataGV1.Rows[i].Cells[5].Value.ToString());
                    }
                    else
                        dataGV2.Rows[i].Cells[47].Value = null;
                }


                connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
        }
        //nhom 1
        private string group1(string chuoi1 )
        {
            if (chuoi1 == "Instant Food and Condiment" || chuoi1 == "Counter Drink- Coffee" || chuoi1 == "Non-food" || chuoi1 == "Snacks and Other Confectionery"
                || chuoi1 == "Instant Noodles" || chuoi1 == "Beverage" || chuoi1 == "Beer" || chuoi1 == "Smaller size beverage"
                || chuoi1 == "Counter Drink- Slurpee" || chuoi1 == "Counter Drink - Coffee & Tea")
                return "E";
            else if (chuoi1 == "Counter Food (Frozen)" || chuoi1 == "Ice-cream and Frozen Food" || chuoi1 == "Smaller Frozen Food")

                return "F";
            else if (chuoi1 == "Chocolates, Candies and Gums" || chuoi1 == "Cigarettes" || chuoi1 == "Wine and Spirits")
                return "W";

            return "S";
        }
        //nhom2
        private string group2(string chuoi1)
        {
            if (chuoi1 == "Instant Food and Condiment" || chuoi1 == "Counter Drink- Coffee" || chuoi1 == "Counter Drink- Slurpee" || chuoi1 == "Counter Drink - Coffee & Tea")
                return "1";
            else if (chuoi1 == "Non-food" || chuoi1 == "Chocolates, Candies and Gums" || chuoi1 == "Wine and Spirits" || chuoi1 == "Cigarettes")
                return "2";
            else if (chuoi1 == "Snacks and Other Confectionery")
                return "3";
            else if (chuoi1 == "Instant Noodles")
                return "4";
            else if (chuoi1 == "Beverage" || chuoi1 == "Smaller size beverage")
                return "5";
            else if (chuoi1 == "Beer")
                return "6";
            else if (chuoi1 == "Counter Food (Frozen)" || chuoi1 == "Ice-cream and Frozen Food"||chuoi1== "Smaller Frozen Food")
                return "R";
            return "F";

        }
        /// Exports the datagridview values to Excel. 
        //private void export2Excel(DataGridView g, string tenTap)
        //{
        //    app obj = new app();
        //    obj.Application.Workbooks.Add(Type.Missing);
        //    obj.Columns.ColumnWidth = 25;
        //    for (int i = 1; i < g.Columns.Count + 1; i++) { obj.Cells[1, i] = g.Columns[i - 1].HeaderText; }
        //    for (int i = 0; i < g.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < g.Columns.Count; j++)
        //        {
        //            if (g.Rows[i].Cells[j].Value != null) { obj.Cells[i + 2, j + 1] = g.Rows[i].Cells[j].Value.ToString(); }
        //        }
        //    }
        //    obj.ActiveWorkbook.SaveCopyAs(tenTap + ".xlsx");
        //    obj.ActiveWorkbook.Saved = true;
        //}
       
        private void btnSave_Click(object sender, EventArgs e)
        {
            
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "ITEM";
            // storing header part in Excel  


            app obj = new app();
            obj.Application.Workbooks.Add(Type.Missing);
            obj.Columns.ColumnWidth = 25;
            for (int i = 1; i < dataGV2.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGV2.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < dataGV2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGV2.Columns.Count; j++)
                {
                    if (dataGV2.Rows[i].Cells[j].Value != null)
                        worksheet.Cells[i + 2, j + 1] = dataGV2.Rows[i].Cells[j].Value.ToString();
                   
                }
            }

           //          save the application
           //SaveFileDialog saveDialog = new SaveFileDialog();
           // saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
           // saveDialog.FilterIndex = 2;
           // if (saveDialog.ShowDialog() == DialogResult.OK)
           // {

           //     export2Excel(dataGV2, saveDialog.FileName);
           //     MessageBox.Show("Export Successful");
           // }
               workbook.SaveAs("C:\\output.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // Exit from the application  
            app.Quit();
        }
    }
}

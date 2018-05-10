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
using Microsoft.VisualBasic.FileIO;
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
            OpenFileDialog openFileDialog = new OpenFileDialog();
            // To list only csv files, we need to add this filter
            openFileDialog.Filter = "|*.csv";
            DialogResult result = openFileDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                try
                {
                    txtPF.Text = openFileDialog.FileName;
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            try
            {

                dataGV1.DataSource = GetDataTableFromCSVFile(txtPF.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Import CSV File", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static DataTable GetDataTableFromCSVFile(string csvfilePath)
        {
            DataTable csvData = new DataTable();
            using (TextFieldParser csvReader = new TextFieldParser(csvfilePath))
            {
                csvReader.SetDelimiters(new string[] { "," });
                csvReader.HasFieldsEnclosedInQuotes = true;

                //Read columns from CSV file, remove this line if columns not exits  
                string[] colFields = csvReader.ReadFields();

                foreach (string column in colFields)
                {
                    DataColumn datecolumn = new DataColumn(column);
                    datecolumn.AllowDBNull = true;
                    csvData.Columns.Add(datecolumn);
                }

                while (!csvReader.EndOfData)
                {
                    string[] fieldData = csvReader.ReadFields();
                    //Making empty value as null
                    for (int i = 0; i < fieldData.Length; i++)
                    {
                        if (fieldData[i] == "")
                        {
                            fieldData[i] = null;
                        }
                    }
                    csvData.Rows.Add(fieldData);
                }
            }
            return csvData;
        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = GetDataTableFromCSVFile(txtPF.Text); //lay du lieu cua datagv1
            try
            {
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
            DataTable dt = (DataTable)dataGV2.DataSource;
            // Get an excel instance
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            // Get a workbook
            Microsoft.Office.Interop.Excel._Workbook wb = excel.Workbooks.Add();

            // Get a worksheet
            Microsoft.Office.Interop.Excel._Worksheet ws = wb.Worksheets.Add();
            ws.Name = "ITEM";

            // Add column names to the first row
           // int col = 1;
            //foreach (DataColumn c in dt.Columns)
            //{
            //    ws.Cells[1, col] = c.ColumnName;
            //    col++;
            //}
            for (int j = 1; j < dataGV2.Columns.Count + 1; j++)
            {
                ws.Cells[1, j] = dataGV2.Columns[j - 1].HeaderText;
            }
            // Create a 2D array with the data from the table
            int i = 0;
            string[,] data = new string[dt.Rows.Count, dt.Columns.Count];
            foreach (DataRow row in dt.Rows)
            {
                int j = 0;
                foreach (DataColumn c in dt.Columns)
                {
                    data[i, j] = row[c].ToString();
                    j++;
                }
                i++;
            }

            // Set the range value to the 2D array
            ws.Range[ws.Cells[2, 1], ws.Cells[dt.Rows.Count + 1, dt.Columns.Count]].value = data;

            // Auto fit columns and rows, show excel, save.. etc
            excel.Columns.AutoFit();
            excel.Rows.AutoFit();
            excel.Visible = true;
        }
    }
}

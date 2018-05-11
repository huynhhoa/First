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
using System.IO;
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
                //them hang o cột đầu tiên
                //foreach (DataRow dr in dt.Rows ) 
                //{
                //    int n = dataGV2.Rows.Add();
                //    dataGV2.Rows[n].Cells[0].Value = dr[0].ToString(); //dr[0] la hàng đầu tiên của datagv1 . datagv2.row[n].celss[0] là cột đầu tiên của dòng đầu tiên.
                //}
                for(int i =0; i <dt.Rows.Count;i++)
                {
                    dataGV2.Rows.Add();
                    dataGV2.Rows[i+8].Cells[0].Value = dataGV1.Rows[i].Cells[0].Value.ToString();
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dataGV2.Rows[i + 8].Cells[1].Value = 99;
                    dataGV2.Rows[i + 8].Cells[2].Value = dataGV1.Rows[i].Cells[1].Value.ToString();
                    dataGV2.Rows[i + 8].Cells[3].Value = "N";
                    dataGV2.Rows[i + 8].Cells[4].Value = "F";
                    dataGV2.Rows[i + 8].Cells[5].Value = "F";
                    dataGV2.Rows[i + 8].Cells[6].Value = "";
                    dataGV2.Rows[i + 8].Cells[7].Value = "A";
                    dataGV2.Rows[i + 8].Cells[8].Value = "";
                    dataGV2.Rows[i + 8].Cells[9].Value = "";
                    dataGV2.Rows[i + 8].Cells[10].Value = "";
                    dataGV2.Rows[i + 8].Cells[11].Value = "";
                    dataGV2.Rows[i + 8].Cells[12].Value = "";
                    dataGV2.Rows[i + 8].Cells[13].Value = "";
                    dataGV2.Rows[i + 8].Cells[14].Value = "";
                    dataGV2.Rows[i + 8].Cells[15].Value = "";
                    dataGV2.Rows[i + 8].Cells[16].Value = "";
                    dataGV2.Rows[i + 8].Cells[17].Value = dataGV1.Rows[i].Cells[10].Value.ToString();
                    dataGV2.Rows[i + 8].Cells[18].Value = "";
                    dataGV2.Rows[i + 8].Cells[19].Value = group1(dataGV1.Rows[i].Cells[37].Value.ToString());
                    dataGV2.Rows[i + 8].Cells[20].Value = group2(dataGV1.Rows[i].Cells[37].Value.ToString());
                    dataGV2.Rows[i + 8].Cells[21].Value = 0;
                    dataGV2.Rows[i + 8].Cells[22].Value = 0;
                    dataGV2.Rows[i + 8].Cells[23].Value = 1;
                    dataGV2.Rows[i + 8].Cells[24].Value = 0;
                    dataGV2.Rows[i + 8].Cells[25].Value = 0;
                    dataGV2.Rows[i + 8].Cells[26].Value = 1;
                    dataGV2.Rows[i + 8].Cells[27].Value = 0;
                    dataGV2.Rows[i + 8].Cells[28].Value = 1;
                    dataGV2.Rows[i + 8].Cells[29].Value = "";
                    dataGV2.Rows[i + 8].Cells[30].Value = "";
                    dataGV2.Rows[i + 8].Cells[31].Value = "";
                    dataGV2.Rows[i + 8].Cells[32].Value = "";
                    dataGV2.Rows[i + 8].Cells[33].Value = "";
                    dataGV2.Rows[i + 8].Cells[34].Value = "";
                    dataGV2.Rows[i + 8].Cells[35].Value = "";
                    dataGV2.Rows[i + 8].Cells[36].Value = "";
                    dataGV2.Rows[i + 8].Cells[37].Value = "";
                    dataGV2.Rows[i + 8].Cells[38].Value = "";
                    dataGV2.Rows[i + 8].Cells[39].Value = "";
                    dataGV2.Rows[i+8].Cells[40].Value = "UPC";
                    dataGV2.Rows[i + 8].Cells[41].Value = dataGV1.Rows[i].Cells[22].Value.ToString();
                    dataGV2.Rows[i + 8].Cells[42].Value = "";
                    dataGV2.Rows[i + 8].Cells[43].Value = "";
                    dataGV2.Rows[i + 8].Cells[44].Value = "";
                    if (dataGV1.Rows[i].Cells[7].Value.ToString() == "")
                        dataGV1.Rows[i].Cells[7].Value = 0;
                    dataGV2.Rows[i + 8].Cells[45].Value = Int32.Parse(dataGV1.Rows[i].Cells[7].Value.ToString());

                    if (int.Parse(dataGV2.Rows[i + 8].Cells[45].Value.ToString()) > 0)
                        dataGV2.Rows[i + 8].Cells[46].Value = 2;
                    else
                        dataGV2.Rows[i + 8].Cells[46].Value = null;

                    if (dataGV1.Rows[i].Cells[6].Value.ToString() == "")
                        dataGV1.Rows[i].Cells[6].Value = 0;

                    if (int.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 2)
                    {
                        dataGV2.Rows[i + 8].Cells[47].Value = int.Parse(dataGV1.Rows[i].Cells[5].Value.ToString()) * 30;
                    }
                    else if (int.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 1)
                    {
                        dataGV2.Rows[i + 8].Cells[47].Value = int.Parse(dataGV1.Rows[i].Cells[5].Value.ToString());
                    }
                    else
                        dataGV2.Rows[i + 8].Cells[47].Value = null;
                    }
                MessageBox.Show("Conversion Successful", "NOTE");

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
        private void copyAlltoClipboard()
        {
            dataGV2.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            dataGV2.MultiSelect = true;
            dataGV2.RowHeadersVisible = false;

            dataGV2.SelectAll();
            DataObject dataObj = dataGV2.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /// Exports the datagridview values to Excel. 
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            { 
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents (*.csv)|*.csv";
            sfd.FileName = "ITEM.csv";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    // Copy DataGridView results to clipboard
                    copyAlltoClipboard();

                    object misValue = System.Reflection.Missing.Value;
                    Excel.Application xlexcel = new Excel.Application();

                    xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                    Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    xlWorkSheet.Name = "ITEM";
                    // Paste clipboard results to worksheet range
                    Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                    // Save the excel file under the captured location from the SaveFileDialog
                    xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlexcel.DisplayAlerts = true;
                    xlWorkBook.Close(true, misValue, misValue);
                    xlexcel.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlexcel);

                    // Clear Clipboard and DataGridView selection
                    Clipboard.Clear();
                    dataGV2.ClearSelection();
                    MessageBox.Show("The File have been Saved", "NOTE", MessageBoxButtons.OK);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            dataGV2.ColumnCount = 55;
            string[] row = new string[] { };
            //dataGV2.Columns[0].Name = "LOGIN ID :";
          //  dataGV2.Columns[1].Name = "NEVSONWH";
            row = new string[] { "LOGIN ID :", "NEVSONWH" };
            dataGV2.Rows.Add(row);
            row = new string[] { "WHS :", "SON" };
            dataGV2.Rows.Add(row);
            row = new string[] { "CUST. CODE :", "SSVN" };
            dataGV2.Rows.Add(row);
            row = new string[] { "OP CODE:", "888" };
            dataGV2.Rows.Add(row);
            row = new string[] { "PASSWORD:", "6868" };
            dataGV2.Rows.Add(row);
            row = new string[] { "REPLY EMAIL :", "nguyenqtan@nittsu.com.hk" };
            dataGV2.Rows.Add(row);
            row = new string[] { };
            dataGV2.Rows.Add(row);
            row = new string[] { "ITEM", "ORIGIN", "DESC", "SRL CTL", "CUSTOMS", "B/F", "HSCODE", "PICK TYPE", "QTY/PKG",
                "VAT","DUTY","LENGTH","WIDTH","HEIGHT","M3","KG","PER PC/PKG","DESC(CUST)","ITEM(SUPPLIER)","UNIT-A","UNIT-A",
                "RACKA:","RACKB:","RACKC:","POSA:","POSB:","POSC:","LVLA:","LVLB:","Currency","Unit price","LENGTH2"
                ,"WIDTH2","HEIGHT2","M32","KG2","CATEGORY 1","CATEGORY 2","CATEGORY 3", "PRODUCT DESC (G)"
            ,"Barcode Kind","Barcode","WPTERM","WPTTYP","SRL OUT","THRSHOLD QTY","REF SORT KEY","MAX QTY OF W/H"};
            dataGV2.Rows.Add(row);
        }
    }
}

﻿using System;
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
using System.Threading;
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
                    try
                    {
                        dataGV1.DataSource = null;
                        dataGV1.Rows.Clear();
                        //do du lieu vao dataGV1
                        dataGV1.DataSource = GetDataTableFromCSVFile(txtPF.Text);
                        //clear dataGV2
                        dataGV2.DataSource = null;
                        dataGV2.Rows.Clear();
                        dataGV2.ColumnCount = 48;
                        string[] row = new string[] { };
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
                        DataTable dt = new DataTable();
                        dt = GetDataTableFromCSVFile(txtPF.Text); //lay du lieu cua datagv1
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            dataGV2.Rows.Add();
                            dataGV2.Rows[i + 8].Cells[0].Value = dataGV1.Rows[i].Cells[0].Value.ToString();
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
                            dataGV2.Rows[i + 8].Cells[40].Value = "UPC";
                            //       dataGV2.Rows[i + 8].Cells[41].Value = dataGV1.Rows[i].Cells[22].Value.ToString();

                            double n;//=double.Parse(dataGV1.Rows[i].Cells[41].Value.ToString());
                            if (double.TryParse(dataGV1.Rows[i].Cells[22].Value.ToString(), out n)) //neu hàng đó là number thì hàng đó bằng n
                            {
                                dataGV2.Rows[i + 8].Cells[41].Value = n;
                            }
                            else //nếu hàng đó là chuỗi thì 
                            {
                                dataGV2.Rows[i + 8].Cells[41].Value = dataGV1.Rows[i].Cells[22].Value.ToString();
                            }
                            dataGV2.Rows[i + 8].Cells[42].Value = "";
                            dataGV2.Rows[i + 8].Cells[43].Value = "";
                            dataGV2.Rows[i + 8].Cells[44].Value = "";

                            dataGV2.Rows[i + 8].Cells[45].Value = dataGV1.Rows[i].Cells[7].Value.ToString();
                            if (dataGV2.Rows[i + 8].Cells[45].Value.ToString() == "")
                                dataGV2.Rows[i + 8].Cells[45].Value = 0;


                            if (double.Parse(dataGV2.Rows[i + 8].Cells[45].Value.ToString()) > 0)
                                dataGV2.Rows[i + 8].Cells[46].Value = 2;
                            else
                                dataGV2.Rows[i + 8].Cells[46].Value = "";

                            if (dataGV1.Rows[i].Cells[6].Value.ToString() == "")
                                dataGV1.Rows[i].Cells[6].Value = 0;

                            if (double.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 2)
                            {
                                dataGV2.Rows[i + 8].Cells[47].Value = double.Parse(dataGV1.Rows[i].Cells[5].Value.ToString()) * 30;
                            }
                            else if (double.Parse(dataGV1.Rows[i].Cells[6].Value.ToString()) == 1)
                            {
                                dataGV2.Rows[i + 8].Cells[47].Value = double.Parse(dataGV1.Rows[i].Cells[5].Value.ToString());
                            }
                            else
                                dataGV2.Rows[i + 8].Cells[47].Value = null;
                        }
                        MessageBox.Show("Conversion Successful", "Sucess");

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
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
        //nhom 1
        private string group1(string chuoi1)
        {
            if (chuoi1 == "Instant Food and Condiment" || chuoi1 == "Counter Drink- Coffee" || chuoi1 == "Non-food" || chuoi1 == "Snacks and Other Confectionery"
                || chuoi1 == "Instant Noodles" || chuoi1 == "Beverage" || chuoi1 == "Beer" || chuoi1 == "Smaller size beverage"
                || chuoi1 == "Counter Drink- Slurpee" || chuoi1 == "Counter Drink - Coffee & Tea")
                return "E";
            else if (chuoi1 == "Counter Food (Frozen)" || chuoi1 == "Ice-cream and Frozen Food" || chuoi1 == "Smaller Frozen Food")

                return "F";
            else if (chuoi1 == "Chocolates, Candies and Gums" || chuoi1 == "Cigarettes" || chuoi1 == "Wine and Spirits")
                return "W";
            else if (chuoi1 == "Store Supply")
                return "S";
            return null;
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
            else if (chuoi1 == "Counter Food (Frozen)" || chuoi1 == "Ice-cream and Frozen Food" || chuoi1 == "Smaller Frozen Food")
                return "R";
            else if (chuoi1 == "Store Supply")
                return "F";
            return null;

        }
        /// Exports the datagridview values to Excel. 
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (dataGV2.Rows.Count != 0)
            {
                try
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "CSV (*.csv)|*.csv";
                    sfd.FileName = "ITEM.csv";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // create one file gridview.csv in writing mode using streamwriter
                        StreamWriter sw = new StreamWriter(sfd.FileName);
                        // iterate through all the rows within the gridview
                        foreach (DataGridViewRow dr in dataGV2.Rows)
                        {
                            // iterate through all colums of specific row
                            for (int i = 0; i < dataGV2.Columns.Count; i++)
                            {
                                // write particular cell to csv file
                                sw.Write(dr.Cells[i].Value);
                                if (i != dataGV2.Columns.Count)
                                {
                                    sw.Write(",");
                                }
                            }
                            // write new line
                            sw.Write(sw.NewLine);
                        }
                        MessageBox.Show("Save Successfully", "Note");
                        // flush from the buffers.
                        sw.Flush();
                        // closes the file
                        sw.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Please Note", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("You do not have any data", "NOTE");

            }

        }
       
       
    }

    
}

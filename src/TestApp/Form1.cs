using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;

namespace TestApp
{
    public partial class Form1 : Form
    {
        private DataSet ds;

        public Form1()
        {
            InitializeComponent();
        }

        /*
        public static void GetValues(DataSet dataset, string sheetName)
        {
            foreach (DataRow row in dataset.Tables[sheetName].Rows)
            {
                foreach (var value in row.ItemArray)
                {
                    Console.WriteLine("{0}, {1}", value, value.GetType());
                }
            }
        }
        */

        private static IList<string> GetTablenames(DataTableCollection tables)
        {
            var tableList = new List<string>();
            foreach (var table in tables)
            {
                tableList.Add(table.ToString());
            }

            return tableList;
        }

        private void Button1Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void Button2Click(object sender, EventArgs e)
        {
            var file = new FileInfo(textBox1.Text);
            using (var stream = new FileStream(textBox1.Text, FileMode.Open))
            {
                IExcelDataReader reader = null;
                if (file.Extension == ".xls")
                {
                   reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (file.Extension == ".xlsx")
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                if (reader == null)
                    return;
                reader.IsFirstRowAsColumnNames = firstRowNamesCheckBox.Checked;
                ds = reader.AsDataSet();

                var tablenames = GetTablenames(ds.Tables);
                sheetCombo.DataSource = tablenames;

                if (tablenames.Count > 0)
                    sheetCombo.SelectedIndex = 0;

                // dataGridView1.DataSource = ds;
                // dataGridView1.DataMember
            }
        }

        private void SelectTable()
        {
            var tablename = sheetCombo.SelectedItem.ToString();

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ds; // dataset
            dataGridView1.DataMember = tablename;

            // GetValues(ds, tablename);
        }

        private void SheetComboSelectedIndexChanged(object sender, EventArgs e)
        {
            SelectTable();
        }
    }
}

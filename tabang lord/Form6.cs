using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tabang_lord
{
    public partial class Form6 : Form
    {
        Workbook book = new Workbook();
        public Form6()
        {
            InitializeComponent();
            DisplayLogs();


        }


        public void DisplayLogs()
        {

            book.LoadFromFile(@"C:\Users\Jojie\OneDrive\Desktop\TECSON\Book25.xlsx"); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[1];
            DataTable dt = new DataTable();
            dt = sheet.ExportDataTable();

            dataGridView1.DataSource = dt;
        }



        private void btnBack_Click(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\Jojie\OneDrive\Desktop\TECSON\Book25.xlsx"); //Change the path to where is the excel locate.
            Worksheet sheet = book.Worksheets[0];
            book.SaveToFile(@"C:\Users\Jojie\OneDrive\Desktop\TECSON\Book25.xlsx");
            this.Hide();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string searchText = txtSearch.Text.Trim().ToLower();

            dataGridView1.ClearSelection();
            dataGridView1.CurrentCell = null;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    string firstName = row.Cells[0].Value?.ToString().ToLower();

                    row.Visible = string.IsNullOrEmpty(searchText) || (firstName != null && firstName.Contains(searchText));
                }
            }
        }
    }
}

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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }
        Form4 form4;
        private void btnLogin_Click(object sender, EventArgs e)
        {
            // Check if BOTH username and password fields are empty
            if (string.IsNullOrWhiteSpace(txtUserName.Text) && string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Please enter both Username and Password.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUserName.Focus();
                return;
            }

            // Check if ONLY username is empty
            if (string.IsNullOrWhiteSpace(txtUserName.Text))
            {
                MessageBox.Show("Please enter your Username.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUserName.Focus();
                return;
            }

            // Check if ONLY password is empty
            if (string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("Please enter your Password.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPassword.Focus();
                return;
            }

            // Proceed to login if both fields are filled
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\TECSON\Book25.xlsx"); // Adjust path as needed
            Worksheet sheet = book.Worksheets[0];

            int row = sheet.LastRow;
            bool log = false;

            for (int i = 2; i <= row; i++) // Start at row 2 to skip header
            {
                string username = sheet.Range[i, 9].Value?.Trim();
                string password = sheet.Range[i, 10].Value?.Trim();

                if (username == txtUserName.Text.Trim() && password == txtPassword.Text.Trim())
                {
                    log = true;
                    admin.Name = sheet.Range[i, 1].Value;
                    form4 = new Form4(admin.Name);
                    form4.pictureBox1.ImageLocation = sheet.Range[i, 12].Value;
                    break;
                }
            }

            if (log)
            {
                MessageBox.Show("Successfully Logged in", "Log in", MessageBoxButtons.OK, MessageBoxIcon.Information);
                mylogs.Log(admin.Name, "Has been logged in");
                form4.Show();
            }
            else
            {
                MessageBox.Show("Invalid Username or Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtUserName_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to Exit?", "Logout Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Log out logic here
                MessageBox.Show("You have been logged out.");
                Application.Exit(); // This will close the application
            }
         
        }
    }
    }


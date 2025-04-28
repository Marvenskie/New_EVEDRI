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

namespace New_EVEDRI
{
    public partial class LogIn: Form
    {
        public LogIn()
        {
            InitializeComponent();
        }

        private void btnLogIn_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx");
            Worksheet SHeet = workbook.Worksheets[0];

            int row = SHeet.Rows.Length;
            bool log = true;

            for (int i = 2; i <= row; i++)
            {
                if (SHeet.Range[i, 6].Value == txtUsername.Text && SHeet.Range[i, 7].Value == txtPassword.Text)
                {
                    log = true; break;
                }
                else
                {
                    log = false;
                }


            }
            if (log == true)
            {
                MessageBox.Show("You have successfully logged in");
                this.Hide();
                AddStud form2 = new AddStud();
                form2.ShowDialog();
               
            }
            else
            {
                MessageBox.Show("Log In Failed");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       
    }
}

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
    public partial class AddStud : Form
    {
        Form3 form3 = new Form3();

        string[] Student = new string[5];
        int i = 0;
        string data = "";
        string gender = "";
        string hobbies = "";
        string favorite = "";
        string saying = "";
        string course = "";
        string Username = "";
        string Password = "";
        public AddStud()
        {
            InitializeComponent();
            btnUpdate.Visible = false;
            lblStatus.Visible = false;
            cmbStatus.Visible = false;

        }



        private void Form2_Load(object sender, EventArgs e)
        {

        }

       

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }





        private void btnActiveStud_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            
            form3.ShowDialog();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtName.Clear();
            radFemale.Checked = false;
            radMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            cmbCourse.SelectedIndex = -1;

            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            data = "";
            gender = "";
            hobbies = "";
            favorite = "";
            saying = "";
            Username = "";
            Password = "";
            course = "";


            if (!string.IsNullOrEmpty(txtName.Text))
            {
                data += txtName.Text;
            }
            if (radFemale.Checked)
            {
                gender += radFemale.Text;
            }
            else if (radMale.Checked)
            {
                gender += radMale.Text;
            }
            if (chkBadminton.Checked)
            {
                hobbies += chkBadminton.Text + ", ";
            }
            if (chkBasketball.Checked)
            {
                hobbies += chkBasketball.Text + ", ";
            }
            if (chkVolleyball.Checked)
            {
                hobbies += chkVolleyball.Text + ", ";
            }
            if (cmbColor.SelectedIndex == 0)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 1)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 2)
            {
                favorite += cmbColor.Text;
            }
            if (cmbCourse.SelectedIndex == 0)
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 1)
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 2)
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 3)
            {
                course += cmbCourse.Text;
            }
            if (!string.IsNullOrEmpty(txtUname.Text))
            {
                Username = txtUname.Text;
            }
            if (!string.IsNullOrEmpty(txtPword.Text))
            {
                Password = txtPword.Text;
            }
            if (!string.IsNullOrEmpty(txtSaying.Text))
            {
                saying += txtSaying.Text;
            }

            Student[i] = data + gender + hobbies + favorite + saying;
            MessageBox.Show("You successfully stored your information", "Successfully", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            int rows = worksheet.Rows.Length + 1;

            worksheet.Range[rows, 1].Value = txtName.Text;
            worksheet.Range[rows, 2].Value = gender;
            worksheet.Range[rows, 3].Value = hobbies;
            worksheet.Range[rows, 4].Value = favorite;
            worksheet.Range[rows, 5].Value = saying;
            worksheet.Range[rows, 6].Value = txtUname.Text;
            worksheet.Range[rows, 7].Value = txtPword.Text;
            worksheet.Range[rows, 9].Value = "Active";

            workbook.SaveToFile(@"C:\Users\Admin\Desktop\New_Excel.xlsx", ExcelVersion.Version2016);

            DataTable dt = worksheet.ExportDataTable();
            form3.dgvData.DataSource = dt;

            txtName.Clear();
            radFemale.Checked = false;
            radMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            cmbCourse.SelectedIndex = -1;

            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
        }
        

        private void btnSubmit_Click_1(object sender, EventArgs e)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx");
            Worksheet sheet = book.Worksheets[0];
            data = "";
            gender = "";
            hobbies = "";
            favorite = "";
            saying = "";
            Username = "";
            Password = "";

            if (!string.IsNullOrEmpty(txtName.Text))
            {
                data += txtName.Text;
            }
            if (radFemale.Checked)
            {
                gender += radFemale.Text;
            }
            else if (radMale.Checked)
            {
                gender += radMale.Text;
            }
            if (chkBadminton.Checked)
            {
                hobbies += chkBadminton.Text + ", ";
            }
            if (chkBasketball.Checked)
            {
                hobbies += chkBasketball.Text + ", ";
            }
            if (chkVolleyball.Checked)
            {
                hobbies += chkVolleyball.Text + ", ";
            }
            if (cmbColor.SelectedIndex == 0)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 1)
            {
                favorite += cmbColor.Text;
            }
            if (cmbColor.SelectedIndex == 2)
            {
                favorite += cmbColor.Text;
            }
            if (!string.IsNullOrEmpty(txtSaying.Text))
            {
                saying += txtSaying.Text;
            }
            if (!string.IsNullOrEmpty(txtUname.Text))
            {
                Username = txtUname.Text;
            }
            if (!string.IsNullOrEmpty(txtPword.Text))
            {
                Password = txtPword.Text;
            }
            if (cmbCourse.SelectedIndex == 0)
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 1) 
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 2)
            {
                course += cmbCourse.Text;
            }
            if (cmbCourse.SelectedIndex == 3)
            {
                course += cmbCourse.Text;
            }


            Student[i] = data + gender + hobbies + favorite + saying;
            MessageBox.Show("You successfully stored your information", "Successfully", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

            int row = sheet.Rows.Length + 1;

            sheet.Range[row, 1].Value = txtName.Text;
            sheet.Range[row, 2].Value = gender;
            sheet.Range[row, 3].Value = hobbies;
            sheet.Range[row, 4].Value = favorite;
            sheet.Range[row, 5].Value = saying;
            sheet.Range[row, 6].Value = txtUname.Text;
            sheet.Range[row, 7].Value = txtPword.Text;
            sheet.Range[row, 8].Value = course;
            sheet.Range[row, 9].Value = "Active";
            book.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx", ExcelVersion.Version2016);

            DataTable dt = sheet.ExportDataTable();
            form3.dgvData.DataSource = dt;

            txtName.Clear();
            radFemale.Checked = false;
            radMale.Checked = false;
            chkBadminton.Checked = false;
            chkBasketball.Checked = false;
            chkVolleyball.Checked = false;
            cmbColor.SelectedIndex = -1;
            cmbCourse.SelectedIndex = -1;
            txtSaying.Clear();
            txtUname.Clear();
            txtPword.Clear();
        }
    }
}

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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }
        public void LoadExcel()
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx");

            Worksheet sheet = workbook.Worksheets[0];
            DataTable dt = sheet.ExportDataTable();
            dgvData.DataSource = dt;
        }
        public void Insert(int Id, string name, string gender, string hobbies, string favcolor, string saying)
        {
            int i = dgvData.Rows.Add(Id, name, gender, hobbies, favcolor, saying);
        }

        private void dgvData_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            AddStud add = new AddStud();    
            add.ShowDialog();
            AddStud addStud = (AddStud)Application.OpenForms["AddStud"];

            int r = dgvData.CurrentCell.RowIndex;
            addStud.label2.Text = r.ToString();

            addStud.txtName.Text = dgvData.Rows[r].Cells[0].Value.ToString();

            string gender = dgvData.Rows[r].Cells[1].Value.ToString();
            if (gender == "Male")
            {
                addStud.radMale.Checked = true;
            }
            else
            {
                addStud.radFemale.Checked = false;
            }

            string hobbies = dgvData.Rows[r].Cells[2].Value.ToString();
            string[] hARRAY = hobbies.Split(',');

            addStud.chkBasketball.Checked = false;
            addStud.chkVolleyball.Checked = false;
            addStud.chkBadminton.Checked = false;

            foreach (string s in hARRAY)
            {
                string hobby = s.Trim();

                if (hobby == "Basketball")
                {
                    addStud.chkBasketball.Checked = true;
                }
                if (hobby == "Volleyball")
                {
                    addStud.chkVolleyball.Checked = true;
                }
                if (hobby == "Badminton")
                {
                    addStud.chkBadminton.Checked = true;
                }

            }

            string favcolor = dgvData.Rows[r].Cells[3].Value.ToString();


            if (favcolor == "Blue")
            {
                addStud.cmbColor.SelectedIndex = 0;
            }
            if (favcolor == "Red")
            {
                addStud.cmbColor.SelectedIndex = 1;
            }
            if (favcolor == "Pink")
            {
                addStud.cmbColor.SelectedIndex = 2;
            }



            string saying = dgvData.Rows[r].Cells[4].Value.ToString();
            string Username = dgvData.Rows[r].Cells[5].Value.ToString();
            string Password = dgvData.Rows[r].Cells[6].Value.ToString();
            string course = dgvData.Rows[r].Cells[7].Value.ToString();
            string status = dgvData.Rows[r].Cells[8].Value.ToString();


            addStud.txtSaying.Text = saying;
            addStud.txtUname.Text = Username;
            addStud.txtPword.Text = Password;
            addStud.cmbCourse.Text = course;



            addStud.btnSubmit.Visible = false;
            addStud.btnUpdate.Visible = true;
            addStud.cmbStatus.Visible = true;
            addStud.lblStatus.Visible = true;
        }

        private void btnAddStud_Click(object sender, EventArgs e)
        {
            this.Hide();
            AddStud addStud = new AddStud();
            addStud.ShowDialog();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            int row = dgvData.CurrentCell.RowIndex + 2;

            worksheet.Range[row, 9].Value = "Inactive";

            workbook.SaveToFile(@"C:\Users\ACT-STUDENT\Desktop\New_Excel (1).xlsx", ExcelVersion.Version2016);
            DataTable dt = worksheet.ExportDataTable();

            dgvData.DataSource = dt;
        }
    }
}

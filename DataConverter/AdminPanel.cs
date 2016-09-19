using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using System.IO;

namespace DataConverter
{
[Serializable]
    public partial class AdminPanel : Form
    {
        public String excel;
        FormAuthorization formAuth;
        FbConnection fbCon;
        bool excelConnection;

        public AdminPanel(FormAuthorization form,FbConnection con)
        {
            formAuth = form;
            InitializeComponent();
            MessageBox.Show("Вы успешно авторизовались как Администратор", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            fbCon = con;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelConnection = false;
            openFileDialog1.ShowDialog();
            excel = openFileDialog1.FileName;
            label6.Text = openFileDialog1.FileName;
            excelConnection = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GPS formGPS = new GPS(fbCon);
            this.Hide();
            formGPS.Show(this);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (excelConnection)
            {
                FormUpload form2 = new FormUpload(fbCon, excel);
                form2.Show(this);
            }
            else
            {
                MessageBox.Show("Проверьте подключение к серверу, и выберете файл для выгрузки", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AdminPanel_FormClosing(object sender, FormClosingEventArgs e)
        {
            formAuth.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FormKeysReport formKeys = new FormKeysReport(fbCon);
            formKeys.Show(this);
            this.Hide();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            FormPlatformReport formReport = new FormPlatformReport(fbCon);
            formReport.Show();
        }

            
    }
}
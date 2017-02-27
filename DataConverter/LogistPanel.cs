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

namespace DataConverter
{
    public partial class LogistPanel : Form
    {
        public String excel;
        FormAuthorization formAuth;
        FbConnection fbCon;
        public bool excelConnection;
        string login;
        public LogistPanel(FormAuthorization form, FbConnection con,string Login)
        {
            InitializeComponent();
            formAuth = form;
            MessageBox.Show("Вы успешно авторизовались как Логист", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            fbCon = con;
            login = Login;
        }

        private void LogistPanel_FormClosing(object sender, FormClosingEventArgs e)
        {
            formAuth.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelConnection = false;
            openFileDialog1.ShowDialog();
            excel = openFileDialog1.FileName;
            label6.Text = openFileDialog1.FileName;
            excelConnection = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FormUpload form2 = new FormUpload(fbCon, excel, excelConnection);
            form2.Show(this); 
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ForDispatcher formForDisp = new ForDispatcher(fbCon);
            formForDisp.Show();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExcelAnalysis formExcelAnalysis = new ExcelAnalysis();
            formExcelAnalysis.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FormKeysReport formKeys = new FormKeysReport(fbCon);
            formKeys.Show(this);
            this.Hide();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GPS formGPS = new GPS(fbCon);
            formGPS.Show(this);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            FormPlatformReport formReport = new FormPlatformReport(fbCon);
            formReport.Show();
        }
    }
}

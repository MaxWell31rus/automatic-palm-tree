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
        FbTransaction fbTrans;
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
            if (fbCon.State == ConnectionState.Closed)
                fbCon.Open();
            fbTrans = fbCon.BeginTransaction();
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

        private void button1_Click(object sender, EventArgs e)
        {
            if(selectCountQueryAccauntsLogin(textBox1.Text)==0)
            {
                insertQueryAccData(selectQueryMaxAccauntId()+1);
            }
            else
            {
                MessageBox.Show("Данный Логин занят", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private int selectCountQueryAccauntsLogin(String id)
        {
            if (fbCon.State == ConnectionState.Closed)
            {
                try
                {
                    fbCon.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Err");
                }
            }
            String selectString = "SELECT COUNT(ID) FROM DC_ACCAUNTS WHERE LOGIN = '" + id + "' GROUP BY ID";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = Convert.ToInt32(obj.ToString());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            return selectResult;
        }

        private int selectQueryMaxAccauntId()
        {
            if (fbCon.State == ConnectionState.Closed)
            {
                try
                {
                    fbCon.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Err");
                }
            }
            String selectString = "SELECT MAX(ID) FROM DC_ACCAUNTS";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                selectResult = Convert.ToInt32(fbComSelect.ExecuteScalar().ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            if (selectResult != 0)
            {
                return selectResult;
            }
            else
            {
                return Const.READ_ERROR;
            }
        }

        private int insertQueryAccData(int id)
        {
            if (fbCon.State == ConnectionState.Closed)
            {
                try
                {
                    fbCon.Open();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Err");
                }
            }
            String insertString = "INSERT INTO DC_ACCAUNTS (ID,LOGIN,PASSWORD,ROLE) VALUES('" + id.ToString() + "','" + textBox1.Text + "','"+textBox2.Text+"','"+comboBox1.Text+"')";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    MessageBox.Show("Добавление новой учетной записи успешно", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            switch(comboBox3.Text)
            {
                case "Диспетчер":
                    {
                        DispatcherPanel formDisp = new DispatcherPanel(this,fbCon);
                        formDisp.Show();
                        break;
                    }
                case "Логист":
                    {
                        LogistPanel formLogist = new LogistPanel(this, fbCon);
                        formLogist.Show();
                        break;
                    }
            }
        }

            
    }
}
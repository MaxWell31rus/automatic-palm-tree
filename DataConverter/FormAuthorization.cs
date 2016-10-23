using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using FirebirdSql.Data.FirebirdClient;


namespace DataConverter
{
    public partial class FormAuthorization : Form
    {
        AdminPanel formAdmin;
        DispatcherPanel formDispatcher;
        LogistPanel formLogist;
        public FbConnection fbCon;
        public FbTransaction fbTrans;
        public FormAuthorization()
        {            
            InitializeComponent();
            label3.Text = "C:\\Soft\\MapXPlus\\DATABASE\\ecotrans_belgorod.GDB";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String role = null;
            role = authDataFinder(textBox1.Text,textBox2.Text);
            if(role != null)
            {
                switch (role)
                {
                    case "Admin":
                        {
                            formAdmin = new AdminPanel(this, fbCon);
                            formAdmin.Show();
                            this.Hide();
                            break;
                        }
                    case "Logist":
                        {
                            formLogist = new LogistPanel(this, fbCon);
                            formLogist.Show();
                            this.Hide();
                            break;                            
                        }
                    case "Dispatcher":
                        {
                            formDispatcher = new DispatcherPanel(this, fbCon);
                            formDispatcher.Show();
                            this.Hide();
                            break;
                        }
                }
                
            }
            else
            {
                MessageBox.Show("Неправильно введена пара логин,пароль", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ConnectionON()
        {
            FbConnectionStringBuilder fbConStr = new FbConnectionStringBuilder();
            fbConStr.Charset = "WIN1251";
            fbConStr.UserID = "SYSDBA";
            fbConStr.Password = "masterkey";
            fbConStr.Database = "localhost:C:\\Users\\Виталик\\Desktop\\работа\\Новая папка (3)\\ecotrans_belgorod.gdb";//"192.168.0.84:C:\\Soft\\MapXPlus\\DATABASE\\ecotrans_belgorod.GDB";
            fbConStr.ServerType = 0;
            try
            {
                fbCon = new FbConnection(fbConStr.ToString());
                fbCon.Open();
                FbDatabaseInfo fbInfo = new FbDatabaseInfo(fbCon);
                MessageBox.Show("Connection Succesfull \nInfo: " + fbInfo.ServerClass + "; " + fbInfo.ServerVersion, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }

        private String authDataFinder(String login, String password)
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
            String selectString = "SELECT ROLE FROM DC_ACCAUNTS WHERE LOGIN = '" + login + "' AND PASSWORD = '"+password+"'";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbTrans = fbCon.BeginTransaction();
            fbComSelect.Transaction = fbTrans;
            String selectResult = null;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = obj.ToString();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
                fbTrans.Commit();
                fbCon.Close();
            }
            return selectResult;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ConnectionON();                    
        }
    }
}

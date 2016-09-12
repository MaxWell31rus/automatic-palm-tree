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
using System.Xml.Serialization;
using System.IO;

namespace DataConverter
{
[Serializable]
    public partial class FormConnection : Form
    {
        public String database;
        public String login;
        public String password;
        public String excel;
        public FbConnection fbCon;
        bool fbConnection = false;
        bool excelConnecion = false;
        FormAuthorization formAuth;

        public FormConnection(FormAuthorization form)
        {
            formAuth = form;
            InitializeComponent();
            XmlSerializer serializer = new XmlSerializer(typeof(List<String>));
            List<String> lStrDeSer = null;
            try
            {
                TextReader fileStream = new StreamReader("IniData\\LoginData.xml");
                lStrDeSer = new List<String>();
                lStrDeSer = (List<String>)serializer.Deserialize(fileStream);
                fileStream.Close();
                if (lStrDeSer != null)
                {
                    login = lStrDeSer[0];
                    password = lStrDeSer[1];
                    database = lStrDeSer[2];
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }           
            textBox1.Text = login;
            textBox2.Text = password;
            label3.Text = database;

        }

        private bool ConnectionON()
        {
            FbConnectionStringBuilder fbConStr = new FbConnectionStringBuilder();
            fbConStr.Charset = "WIN1251";
            fbConStr.UserID = login;
            fbConStr.Password = password;
            fbConStr.Database = "185.5.17.46:C:\\Soft\\MapXPlus\\DATABASE\\ecotrans_belgorod.GDB";//database;
            fbConStr.ServerType = 0;
            try
            {
                fbCon = new FbConnection(fbConStr.ToString());
                fbCon.Open();
                fbConnection = true;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message,"Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                fbConnection = false;
                this.Close();
            }
            return fbConnection;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            database = openFileDialog1.FileName;
            label3.Text = openFileDialog1.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            login = textBox1.Text;
            password = textBox2.Text;
            ConnectionON();
            FbDatabaseInfo fbInfo = new FbDatabaseInfo(fbCon);
            MessageBox.Show("Connection Succesfull \nInfo: " + fbInfo.ServerClass + "; " + fbInfo.ServerVersion,"Success",MessageBoxButtons.OK,MessageBoxIcon.Information);
            
            if (checkBox1.Checked == true)
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<String>));
                TextWriter fileStream = new StreamWriter("LoginData.xml");
                List<String> lStrSer = new List<String>();
                lStrSer.Add(login);
                lStrSer.Add(password);
                lStrSer.Add(database);
                serializer.Serialize(fileStream, lStrSer);
                fileStream.Close();
            }
            if(checkBox2.Checked == true && checkBox1.Checked == false)
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<String>));
                TextWriter fileStream = new StreamWriter("LoginData.xml");
                List<String> lStrSer = new List<String>();
                lStrSer.Add("");
                lStrSer.Add("");
                lStrSer.Add("");
                serializer.Serialize(fileStream, lStrSer);
                fileStream.Close();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelConnecion = false;
            openFileDialog1.ShowDialog();
            excel = openFileDialog1.FileName;
            label6.Text = openFileDialog1.FileName;
            excelConnecion = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GPS formGPS = new GPS(fbCon);
            this.Hide();
            formGPS.Show(this);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (fbConnection && excelConnecion)
            {
                Form2 form2 = new Form2(fbCon, excel);
                form2.Show(this);
            }
            else
            {
                MessageBox.Show("Проверьте подключение к серверу, и выберете файл для выгрузки", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FormConnection_FormClosing(object sender, FormClosingEventArgs e)
        {
            formAuth.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Keys formKeys = new Keys();
            formKeys.Show(this);
            this.Hide();
        }

            
    }
}
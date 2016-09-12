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

namespace DataConverter
{
    public partial class FormAuthorization : Form
    {
        private String stringAuthData = "192.168.0.84:C:\\Soft\\MapXPlus\\DATABASE\\authdata";
        FormConnection formCon;
        public FormAuthorization()
        {            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //TextReader streamAuthData = new StreamReader(stringAuthData);
            if(String.Compare(textBox1.Text,"")==0 && String.Compare(textBox2.Text,"")==0)
            {
                formCon = new FormConnection(this);
                formCon.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Неправильно введена пара логин,пароль", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

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
        AdminPanel formAdmin;
        FormAuthorization formAuth;
        FbConnection fbCon;
        bool bAdmIn = false;
        public LogistPanel(FormAuthorization form, FbConnection con)
        {
            InitializeComponent();
            formAuth = form;
            MessageBox.Show("Вы успешно авторизовались как Логист", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            fbCon = con;
        }

        public LogistPanel(AdminPanel form, FbConnection con)
        {
            InitializeComponent();
            formAdmin = form;
            MessageBox.Show("Вы успешно авторизовались как Логист", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            fbCon = con;
            bAdmIn = true;
        }

        private void LogistPanel_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(!bAdmIn)
            formAuth.Close();
        }
    }
}

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
        FormAuthorization formAuth;
        FbConnection fbCon;
        public LogistPanel(FormAuthorization form, FbConnection con)
        {
            InitializeComponent();
            formAuth = form;
            MessageBox.Show("Вы успешно авторизовались как Логист", "Авторизация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            fbCon = con;
        }

        private void LogistPanel_FormClosing(object sender, FormClosingEventArgs e)
        {
            formAuth.Close();
        }
    }
}

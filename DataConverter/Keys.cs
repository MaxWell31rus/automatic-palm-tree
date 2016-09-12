using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataConverter
{
    public partial class Keys : Form
    {
        public Keys()
        {
            InitializeComponent();
        }

        private void Keys_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Show();
        }
    }
}

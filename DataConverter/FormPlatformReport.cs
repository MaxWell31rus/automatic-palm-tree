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
    public partial class FormPlatformReport : Form
    {
        FbConnection fbCon;
        FbTransaction fbTrans;
        public FormPlatformReport(FbConnection con)
        {
            InitializeComponent();
            fbCon = con;
            if(fbCon.State == ConnectionState.Closed)
            fbCon.Open();
            fbTrans = fbCon.BeginTransaction();
        }

        private void FormPlatformReport_Load(object sender, EventArgs e)
        {
            selectPlatformWithoutGroup();
        }

        private void selectPlatformWithoutGroup()
        {
            List<List<String>> listPlatforms = new List<List<String>>();
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
            String selectString = "select unistring,name,fullname,adress,gps_x_mobile,gps_y_mobile from clients where (id not in (select client_id from groupclientsstring)) and (gps_x_mobile is not null)";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            FbDataReader selectResult = null;
            try
            {
                selectResult = fbComSelect.ExecuteReader();
                while (selectResult.Read())
                {
                    listPlatforms.Add(new List<String>());
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(0)));
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(1)));
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(2)));
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(3)));
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(4)));
                    listPlatforms[listPlatforms.Count - 1].Add(String.Copy(selectResult.GetString(5)));                    
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
                selectResult.Dispose();
                fbTrans.Commit();
                fbTrans.Dispose();
                fbCon.Close();
            }
            for (int i = 0; i < listPlatforms.Count;i++ )
            {
                dataGridView1.Rows.Add();
                for(int j=0;j<listPlatforms[i].Count;j++)
                {
                    dataGridView1[j,i].Value = listPlatforms[i][j];
                }
            }
        }
    }
}

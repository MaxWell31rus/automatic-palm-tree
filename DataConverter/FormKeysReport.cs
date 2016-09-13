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
    public partial class FormKeysReport : Form
    {
        FbConnection fbCon;
        FbTransaction fbTrans;
        public FormKeysReport(FbConnection con)
        {
            InitializeComponent();
            fbCon = con;
            if (fbCon.State == ConnectionState.Closed)
                fbCon.Open();
            fbTrans = fbCon.BeginTransaction();
        }



        private void Keys_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Show();
        }


        private void selectKeys()
        {
            List<List<String>> listKeys = new List<List<String>>();
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
            String selectString = "select " +
                                  "c.unistring as code, " +
                                  "cr.nnumber as car, " +
                                  "k.key_number as key_number, " +
                                  "p.name as driver, " +
                                  "t.docdate as d_date, " +
                                  "c.name as CPName," +
                                  "c.fullname as client," +
                                  "c.adress as adress, " +
                                  "g.name as schedule " +
                                  "from travelsheet_model  tm " +
                                    " left join travelsheet t on t.travelsheet_model_id=tm.id" +
                                    " left join travelsheetstring ts on ts.travel_id=t.id" +
                                    " left join cars cr on cr.id=t.car_id" +
                                    " left join people p on p.id=cr.voditel_id" +
                                    " left join docs d on d.id=ts.doc_id" +
                                    " left join clients c on c.id=ts.client_id" +
                                    " left join variables_string v on v.client_id = c.id" +
                                    " left join groupclientsstring gs on gs.client_id=c.id" +
                                    " left join groupclients g on g.id=gs.groupclients_id" +
                                    " join keys k on c.unistring = k.client_unistring" +
                                  " where tm.model_name_id is not null and t.status = '5' and cr.nnumber = '" + textBox1.Text.ToUpper() + "' and t.docdate >= '" + dateTimePicker1.Value.ToShortDateString() + "' and t.docdate <= '" + dateTimePicker2.Value.ToShortDateString() + "' and v.variables_id = '61'";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            FbDataReader selectResult = null;
            try
            {
                selectResult = fbComSelect.ExecuteReader();
                while (selectResult.Read())
                {
                    listKeys.Add(new List<String>());
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(0)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(1)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(2)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(3)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(4)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(5)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(6)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(7)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(8)));
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
                fbTrans.Dispose();
                fbCon.Close();
            }
            for (int i = 0; i < listKeys.Count; i++)
            {
                dataGridView1.Rows.Add();
                for (int j = 0; j < listKeys[i].Count; j++)
                {
                    dataGridView1[j, i].Value = listKeys[i][j];
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            selectKeys();
        }
    }
}

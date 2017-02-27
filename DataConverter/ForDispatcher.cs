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
    public partial class ForDispatcher : Form
    {
        FbConnection fbCon;
        FbTransaction fbTrans;
        public ForDispatcher(FbConnection con)
        {
            InitializeComponent();
            fbCon = con;
            fbCon.Open();
            fbTrans = fbCon.BeginTransaction();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            selectKeys();
        }

        private void selectKeys()
        {
            int countAll = 0, count = 0;
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
            String selectString = "select " +                //2 3 4 7-
        "c.unistring as code, " +
        "cr.nnumber as car, " +
        "p.name as driver, " +
        "t.docdate as d_date, " +
        "c.name as CPName, " +
        "c.adress as adress, " +
        "c.fullname as client, " +
        "v.value1 as container, " +
        "d.obem as volume, " +
        "ts.reason_time as fact_time, " +
        "g.name as schedule, " +
        "ts.reason_id as status " +
        "from travelsheet_model  tm " +
         "left join travelsheet t on t.travelsheet_model_id=tm.id " +
         "left join travelsheetstring ts on ts.travel_id=t.id " +
         "left join cars cr on cr.id=t.car_id " +
         "left join people p on p.id=cr.voditel_id " +
         "left join docs d on d.id=ts.doc_id " +
         "left join clients c on c.id=ts.client_id " +
         "left join variables_string v on v.client_id = c.id " +
         "left join groupclientsstring gs on gs.client_id=c.id " +
         "left join groupclients g on g.id=gs.groupclients_id " +
        "where t.status = 5 and t.docdate>='" + dateTimePicker1.Value.ToShortDateString() + "' and t.docdate<='" + dateTimePicker1.Value.ToShortDateString() + "' and v.variables_id = '61' and tm.model_name_id is not null and cr.nnumber='" + textBox1.Text.ToUpper() + "'";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            FbDataReader selectResult = null;
            try
            {
                selectResult = fbComSelect.ExecuteReader();
                while (selectResult.Read())
                {
                    listKeys.Add(new List<String>());
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(0)));  //2 3 4 7-
                    //listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(1)));
                    // listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(2)));
                    // listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(3)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(4)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(5)));
                    //listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(6)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(7)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(8)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(9)));
                    listKeys[listKeys.Count - 1].Add(String.Copy(selectResult.GetString(10)));
                    if (String.Compare(selectResult.GetString(11), "1") == 0)
                        listKeys[listKeys.Count - 1].Add("Забрана");
                    else if (String.Compare(selectResult.GetString(11), "") == 0)
                        listKeys[listKeys.Count - 1].Add("Нет статуса");
                    else
                        listKeys[listKeys.Count - 1].Add("Не забрана");
                    countAll++;
                    if (String.Compare(listKeys[listKeys.Count - 1][7], "Забрана") == 0)
                        count++;
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
            }
            for (int i = 0; i < listKeys.Count; i++)
            {
                dataGridView1.Rows.Add();
                for (int j = 0; j < listKeys[i].Count; j++)
                {
                    dataGridView1[j, i].Value = listKeys[i][j];
                }
            }
            label5.Text = countAll.ToString();
            label6.Text = count.ToString();
        }

        private void ForDispatcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                fbTrans.Commit();
                fbCon.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                fbTrans.Dispose();
            }
        }



    }
}

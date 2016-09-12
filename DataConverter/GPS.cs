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
    public partial class GPS : Form
    {
        FbConnection fbCon = null;
        FbTransaction fbTrans;
        public GPS(FbConnection paramFbcon)
        {
            fbCon = paramFbcon;
            InitializeComponent();
            fbTrans = fbCon.BeginTransaction();
        }

        private int selectQueryClientsId(String unistring)
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
            String selectString = "SELECT ID FROM CLIENTS WHERE UNISTRING ='"+ unistring +"'";
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


        private void selectQueryGPS(out string X,out string Y,String id)
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
            String selectString = "select gps_x_mobile ,gps_y_mobile from clients where id = '"+ id +"'";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            FbDataReader selectResult = null;
            X = "";
            Y = "";
            try
            {
                selectResult = fbComSelect.ExecuteReader();
                while (selectResult.Read())
                {
                    Y = String.Copy(selectResult.GetString(0));
                    X = String.Copy(selectResult.GetString(1));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
                fbCon.Close();
            }           
        }

        private int UpdateQueryCLIENTS(String y,String x, String id)
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
            String insertString = "UPDATE CLIENTS SET GPS_X = '" + x + "',GPS_Y = '" + y + "',GPS_X_MOBILE = '" + x + "',GPS_Y_MOBILE = '" + y + "',DIFF_GPS_DIST = '1' WHERE ID = '" + id + "'";
            FbCommand fbComUpdate = new FbCommand(insertString, fbCon);
            fbComUpdate.Transaction = fbTrans;
            int insRes = 0;
            bool close = false;
            try
            {
                insRes = fbComUpdate.ExecuteNonQuery();
                if (insRes == 0)
                {
                    MessageBox.Show("Andrey ti pidoras nepravilno vvel suka");
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                close = true;
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                fbComUpdate.Dispose();
                fbCon.Close();
                if(close)
                {
                    this.Close();
                }
            }
            if (insRes != 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String strX = "";
            String strY = "";
            int id = 0;
            if (checkBox2.Checked == false)
            {
                if (textBox1.Text != "" && textBox2.Text != "" && textBox3.Text != "")
                {
                    if (checkBox1.Checked == true)
                    {
                        double x = Convert.ToDouble(textBox4.Text) / 60 + Convert.ToDouble(textBox2.Text);
                        double y = Convert.ToDouble(textBox5.Text) / 60 + Convert.ToDouble(textBox3.Text);
                        strX = String.Copy(x.ToString());
                        strY = String.Copy(y.ToString());
                    }
                    else
                    {
                        strX = String.Copy(textBox2.Text);
                        strY = String.Copy(textBox3.Text);
                    }
                    id = selectQueryClientsId(textBox1.Text);
                    if (UpdateQueryCLIENTS(strX, strY, id.ToString()) == 0)
                        MessageBox.Show("Err");
                }
                else
                {
                    MessageBox.Show("Не все поля заполнены");
                }
            }
            else
            {
                id = selectQueryClientsId(textBox1.Text);
                selectQueryGPS(out strX,out strY, id.ToString());
                id = selectQueryClientsId(textBox6.Text);
                if (UpdateQueryCLIENTS(strX, strY, id.ToString()) == 0)
                    MessageBox.Show("Err");
            }
            
        }

        private void GPS_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Show();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}

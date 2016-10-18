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
    public partial class FormUpload : Form
    {
        List<List<String>> listTimes = new List<List<String>>();
        FbTransaction fbTrans = null;
        FbConnection fbCon = null;
        String excel = null;
        List<String> lDuplicate = new List<String>();
        List<String> lGroupsClients = new List<String>();
        Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        Microsoft.Office.Interop.Excel.Application ObjExcel;
        Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        bool excelConnection;

        public FormUpload(FbConnection FbCon, String Excel,bool b)
        {
            excelConnection = b;
            fbCon = FbCon;
            fbCon.Open();
            fbTrans = fbCon.BeginTransaction();
            excel = Excel;
            InitializeComponent();
            if (b)
            {
                ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                ObjWorkBook = ObjExcel.Workbooks.Open(excel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            }
        }

        private int excelDataToList(int id, List<Data> lData, Data data)
        {
            int iRet = data.dataReadFromExcel(ObjWorkSheet, id, lData, lDuplicate, lGroupsClients);
            if (iRet == Const.READ_ERROR)
            {
                return -1;
            }
            else if (iRet == Const.READ_SUCCESS)
            {
                id++;
            }
            else if (iRet == Const.READ_ABORT)
            {
                id++;
                id = excelDataToList(id, lData, data);
            }
            return id;
        }

        // получение 2 записи из екселя

        private int InsertQueryGROUPCLIENTS(int id, String groupClient)
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
            String insertString;
            if (Const.SIDE.Contains(groupClient))
            {
                insertString = "INSERT INTO GROUPCLIENTS (ID,NAME,REGION_TYPE_ID,USED_IN_CALC) VALUES('" + id.ToString() + "','" + groupClient + "','5','1')";
            }
            else
            {
                insertString = "INSERT INTO GROUPCLIENTS (ID,NAME,REGION_TYPE_ID,USED_IN_CALC) VALUES('" + id.ToString() + "','" + groupClient + "','5','1')";
            }
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int InsertQueryGROUPCLIENTSSTRING(int clientId, Data values)
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
            int groupId = selectQueryGroupClientsId(values.CPGroupClient);
            String insertString = "INSERT INTO GROUPCLIENTSSTRING (CLIENT_ID,GROUPCLIENTS_ID) VALUES('" + clientId.ToString() + "','" + groupId.ToString() + "')";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int UpdateQueryGROUPCLIENTSSTRING(int clientId, Data values)
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
            int groupId = selectQueryGroupClientsId(values.CPGroupClient);
            String insertString = "UPDATE GROUPCLIENTSSTRING SET GROUPCLIENTS_ID = '" + groupId.ToString() + "' WHERE CLIENT_ID = '" + clientId.ToString() + "'";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes != 0)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes != 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int InsertQueryDAYMONTH(String dayMonth, int id, String used)
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
            String insertString = "INSERT INTO CLIENTS_DAYMONTH_SCHEDULE (CLIENT_ID,DAYMONTH,GRUZ_TYPE_ID,USED) VALUES('" + id.ToString() + "','" + dayMonth + "','2','" + used +"')";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fbTrans.Rollback();
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int deleteQueryDAYMONTH(int id)
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
            String insertString = "DELETE FROM CLIENTS_DAYMONTH_SCHEDULE WHERE CLIENT_ID = '" + id.ToString() + "'";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes != 0)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes != 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int delete(string id)
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
            String insertString = "DELETE FROM GROUPCLIENTSSTRING WHERE GROUPCLIENTS_ID = '"+id+"'";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch(Exception)
            {
                fbTrans.Rollback();
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            return Const.READ_SUCCESS;
        }

        private int deleteQuerySCHEDULE(int id)
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
            String insertString = "DELETE FROM CLIENTS_DELIVERY_SCHEDULE WHERE CLIENT_ID = '" + id.ToString() + "'";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes != 0)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                this.Close();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes != 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int InsertQuerySCHEDULE(String shablonesId, String nmbOfWeek, int id)
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
            String insertString = "INSERT INTO CLIENTS_DELIVERY_SCHEDULE (CLIENT_ID,SHABLONES_ID,NUMBER_OF_WEEK,USED,GRUZ_TYPE_ID) VALUES('" + id.ToString() + "','" + shablonesId + "','" + nmbOfWeek + "','1', '2')";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int InsertQueryCLIENTS(Data values, int id)
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
            String insertString = "INSERT INTO CLIENTS (ID,UNISTRING,NAME,FULLNAME,ADRESS,COMMENTS,UPLIMTIME0,DNLIMTIME0,USED,TIME_WAIT, CLIENTS_TYPE_ID) VALUES('" + id.ToString() + "','" + values.CPCode + "','" + values.CPName + "','" + values.CPOwner + "','" + values.CPAdress + "','" + values.CPComment + "','" + values.CPTimelineB + "','" + values.CPTimelineE + "','1','1','" + values.CPStatus + "')";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
            }
            insertString = "INSERT INTO VARIABLES_STRING (CLIENT_ID,VARIABLES_ID,VALUE1) VALUES('" + id.ToString() + "','61','" + values.CPConCount + "')";
            fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            try
            {
                insRes &= fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                this.Close();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                fbComInsert.Dispose();
            }
            insertString = "INSERT INTO VARIABLES_STRING (CLIENT_ID,VARIABLES_ID,VALUE1) VALUES('" + id.ToString() + "','62','" + values.CPType + "')";
            fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            try
            {
                insRes &= fbComInsert.ExecuteNonQuery();
                if (insRes == 1)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            if (insRes == 1)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int UpdateQueryCLIENTS(Data values, int id)
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
            String insertString = "UPDATE CLIENTS SET COMMENTS = '" + values.CPComment + "',UPLIMTIME0 = '" + values.CPTimelineB + "',DNLIMTIME0 = '" + values.CPTimelineE + "',CLIENTS_TYPE_ID = '" + values.CPStatus + "',adress = '"+ values.CPAdress+"' WHERE ID = '" + id.ToString() + "'";
            FbCommand fbComUpdate = new FbCommand(insertString, fbCon);
            fbComUpdate.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComUpdate.ExecuteNonQuery();
                if (insRes != 0)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                fbTrans.Rollback();
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComUpdate.Dispose();
                fbCon.Close();
            }
            
            int count = selectCountQueryVariablesStringID(id.ToString());
            if (count != 0)
                insertString = "UPDATE VARIABLES_STRING SET VALUE1 = '" + values.CPConCount + "' WHERE (CLIENT_ID = '" + id.ToString() + "' AND VARIABLES_ID = '61')";
            else
                insertString = "INSERT INTO VARIABLES_STRING (CLIENT_ID,VARIABLES_ID,VALUE1) VALUES('" + id.ToString() + "','61','" + values.CPConCount + "')";
            fbComUpdate = new FbCommand(insertString, fbCon);
            fbComUpdate.Transaction = fbTrans;
            try
            {
                insRes &= fbComUpdate.ExecuteNonQuery();
                if (insRes == 1)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComUpdate.Dispose();
            }
            if (count != 0)
                insertString = "UPDATE VARIABLES_STRING SET VALUE1 = '" + values.CPType + "' WHERE CLIENT_ID = '" + id.ToString() + "' AND VARIABLES_ID = '62'";
            else
                insertString = "INSERT INTO VARIABLES_STRING (CLIENT_ID,VARIABLES_ID,VALUE1) VALUES('" + id.ToString() + "','62','" + values.CPType + "')";
            fbComUpdate = new FbCommand(insertString, fbCon);
            fbComUpdate.Transaction = fbTrans;
            try
            {
                insRes &= fbComUpdate.ExecuteNonQuery();
                if (insRes != 0)
                {
                    //  MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Insert Error:" + e.Message + "\nЗапись приостановлена", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                fbComUpdate.Dispose();
                fbCon.Close();
            }              
            if (insRes != 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }

        private int selectQueryMaxClientsId()
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
            String selectString = "SELECT MAX(ID) FROM CLIENTS";
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
            String selectString = "SELECT MAX(ID) FROM CLIENTS WHERE UNISTRING = '" + unistring + "'";
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

        private int selectQueryGroupClientsId(String name)
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
            String selectString = "SELECT ID FROM GROUPCLIENTS WHERE NAME = '" + name + "'";
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

        private int selectQueryMaxGroupClientsId()
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
            String selectString = "SELECT MAX(ID) FROM GROUPCLIENTS";
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

        private int selectCountQueryClientsId(String id)
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
            String selectString = "SELECT COUNT(UNISTRING) FROM CLIENTS WHERE UNISTRING = '" + id + "' GROUP BY UNISTRING";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = Convert.ToInt32(obj.ToString());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            return selectResult;
        }

        private int selectCountQueryGroupClientsId(String id)
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
            String selectString = "SELECT COUNT(ID) FROM GROUPCLIENTS WHERE NAME = '" + id + "' GROUP BY ID";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = Convert.ToInt32(obj.ToString());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            return selectResult;
        }

        private int selectCountQueryGroupClientsStringId(String id)
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
            String selectString = "SELECT COUNT(ID) FROM GROUPCLIENTSSTRING WHERE CLIENT_ID = '" + id + "' GROUP BY ID";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = Convert.ToInt32(obj.ToString());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            return selectResult;
        }

        private int selectCountQueryVariablesStringID(String id)
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
            String selectString = "SELECT COUNT(ID) FROM VARIABLES_STRING WHERE CLIENT_ID = '" + id + "' GROUP BY ID";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            int selectResult = 0;
            try
            {
                object obj = fbComSelect.ExecuteScalar();
                if (obj != null)
                {
                    selectResult = Convert.ToInt32(obj.ToString());
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                fbComSelect.Dispose();
            }
            return selectResult;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (excelConnection)
            {
                int id = selectQueryMaxClientsId() + 1;
                int i = 10;
                List<Data> lData = new List<Data>();
                Data tmp;
                if (checkBox1.Checked == true)
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(List<Data>));
                    TextReader fileStream = new StreamReader("Backup.xml");
                    lData = (List<Data>)serializer.Deserialize(fileStream);
                    fileStream.Close();
                }
                else
                {
                    while (true)
                    {
                        tmp = new Data();
                        i = excelDataToList(i, lData, tmp);
                        if (i != -1)
                        {
                            lData.Add(tmp);
                            label1.Text = (i - 10).ToString();
                        }
                        else
                        {
                            break;
                        }
                    }
                    XmlSerializer serializer = new XmlSerializer(typeof(List<Data>));
                    TextWriter fileStream = new StreamWriter("Backup.xml");
                    serializer.Serialize(fileStream, lData);
                    fileStream.Close();
                }
                int groupId = selectQueryMaxGroupClientsId() + 1;
                for (i = 0; i < lGroupsClients.Count; i++)
                {
                    if (selectCountQueryGroupClientsId(lGroupsClients[i]) == 0)
                    {
                        if (InsertQueryGROUPCLIENTS(groupId + i, lGroupsClients[i]) == Const.READ_ERROR)
                            return;
                    }
                    else
                    {
                        int x = selectQueryGroupClientsId(lGroupsClients[i]);
                        delete(x.ToString());
                    }
                }
                for (i = 0; i < lData.Count; i++)
                {
                    if (selectCountQueryClientsId(lData[i].CPCode) == 0)
                    {
                        if (InsertQueryCLIENTS(lData[i], id + i) == Const.READ_ERROR)
                            break;
                        switch (lData[i].CPSchedule[0])
                        {
                            case "1":
                                {
                                    for (int j = 1; j < lData[i].CPSchedule.Count; j += 2)
                                    {
                                        if (InsertQuerySCHEDULE(Data.convertWeekToNmb(lData[i].CPSchedule[j + 1]), lData[i].CPSchedule[j], id + i) == Const.READ_ERROR)
                                        {
                                            i = lData.Count;
                                            break;
                                        }
                                    }
                                    break;
                                }
                            case "2":
                                {
                                    for (int k = 0; k < Const.NMB_WEEK.Count; k++)
                                    {
                                        for (int j = 1; j < lData[i].CPSchedule.Count; j++)
                                        {
                                            if (InsertQuerySCHEDULE(Data.convertWeekToNmb(lData[i].CPSchedule[j]), Const.NMB_WEEK[k], id + i) == Const.READ_ERROR)
                                            {
                                                i = lData.Count;
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            case "3":
                                {
                                    for (int j = 1; j < lData[i].CPSchedule.Count; j++)
                                    {
                                        if (InsertQueryDAYMONTH(lData[i].CPSchedule[j], id + i, lData[i].CPUsed) == Const.READ_ERROR)
                                        {
                                            i = lData.Count;
                                            break;
                                        }
                                    }
                                    break;
                                }
                        }
                        /*             dsadasdsadasdasdasd                     */
                        if (InsertQueryGROUPCLIENTSSTRING(i + id, lData[i]) == Const.READ_ERROR)
                            break;
                    }
                    else
                    {
                        int updId = selectQueryClientsId(lData[i].CPCode);
                        if (UpdateQueryCLIENTS(lData[i], updId) == Const.READ_ERROR)
                            break;
                        switch (lData[i].CPSchedule[0])
                        {
                            case "1":
                                {
                                    deleteQuerySCHEDULE(updId);
                                    for (int j = 1; j < lData[i].CPSchedule.Count; j += 2)
                                    {
                                        if (InsertQuerySCHEDULE(Data.convertWeekToNmb(lData[i].CPSchedule[j + 1]), lData[i].CPSchedule[j], updId) == Const.READ_ERROR)
                                        {
                                            i = lData.Count;
                                            break;
                                        }
                                    }
                                    break;
                                }
                            case "2":
                                {
                                    deleteQuerySCHEDULE(updId);
                                    for (int k = 0; k < Const.NMB_WEEK.Count; k++)
                                    {
                                        for (int j = 1; j < lData[i].CPSchedule.Count; j++)
                                        {

                                            if (InsertQuerySCHEDULE(Data.convertWeekToNmb(lData[i].CPSchedule[j]), Const.NMB_WEEK[k], updId) == Const.READ_ERROR)
                                            {
                                                i = lData.Count;
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            case "3":
                                {
                                    deleteQueryDAYMONTH(updId);
                                    for (int j = 1; j < lData[i].CPSchedule.Count; j++)
                                    {
                                        if (InsertQueryDAYMONTH(lData[i].CPSchedule[j], updId, lData[i].CPUsed) == Const.READ_ERROR)
                                        {
                                            i = lData.Count;
                                            break;
                                        }
                                    }
                                    break;
                                }
                        }
                        int cnt = selectCountQueryGroupClientsStringId(updId.ToString());
                        if (cnt != 0)
                        {
                            if (UpdateQueryGROUPCLIENTSSTRING(updId, lData[i]) == Const.READ_ERROR)
                                break;
                        }
                        else
                        {
                            if (InsertQueryGROUPCLIENTSSTRING(updId, lData[i]) == Const.READ_ERROR)
                                break;
                        }

                    }
                    label2.Text = (i + 1).ToString() + "/" + lData.Count.ToString();
                    this.Refresh();
                }
            }           
            else
            {
                MessageBox.Show("Проверьте подключение к серверу, и выберете файл для выгрузки", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {

            try
            {
                ObjWorkBook.Close();
                ObjExcel.Quit();
                fbTrans.Commit();
                fbCon.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            deleteQueryGroupClientsString();
        }

        private int deleteQueryGroupClientsString()
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
            String insertString = "delete from groupclientsstring where id in( select groupclientsstring.id from groupclientsstring left join groupclients on groupclients.id = groupclientsstring.groupclients_id where groupclients.region_type_id = 5 and groupclientsstring.groupclients_id <> 163 )";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                // MessageBox.Show("Success", "Insert Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception)
            {
                fbTrans.Rollback();
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            return Const.READ_SUCCESS;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            selectTimes();
            for (int i = 0; i < listTimes.Count; i++)
            {
                insertTimes((Convert.ToDateTime(listTimes[i][1]).AddMinutes(-30)).ToString(), (Convert.ToDateTime(listTimes[i][1]).AddMinutes(30)).ToString(), listTimes[i][0]);
            }
        }

        private void selectTimes()
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
            String selectString = "select " +
                                  "c.id as code, " +
                                  "ts.reason_time as fact_time " +
                                  "from travelsheet_model  tm " +
                                    " left join travelsheet t on t.travelsheet_model_id=tm.id" +
                                    " left join travelsheetstring ts on ts.travel_id=t.id" +                                   
                                    " left join clients c on c.id=ts.client_id" +
                                    " left join variables_string v on v.client_id = c.id" +
                                  " where tm.model_name_id is not null and t.status = '5' and t.docdate >= '" + dateTimePicker1.Value.ToShortDateString() + "' and t.docdate <= '" + dateTimePicker1.Value.ToShortDateString() + "' and v.variables_id = '61'and ts.reason_time is not null";
            FbCommand fbComSelect = new FbCommand(selectString, fbCon);
            fbComSelect.Transaction = fbTrans;
            FbDataReader selectResult = null;
            try
            {
                selectResult = fbComSelect.ExecuteReader();
                while (selectResult.Read())
                {
                    listTimes.Add(new List<String>());
                    listTimes[listTimes.Count - 1].Add(String.Copy(selectResult.GetString(0)));
                    listTimes[listTimes.Count - 1].Add(String.Copy(selectResult.GetString(1)));                   
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
        }

        private int insertTimes(string time1,string time2,string id)
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
            String insertString = "update clients set uplimtime0='" + Convert.ToDateTime(time1).ToShortTimeString() + "',dnlimtime0='" + Convert.ToDateTime(time2).ToShortTimeString() + "' where id = '" + id + "'";
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();
                
            }
            catch (Exception e)
            {
                MessageBox.Show("Success", e.Message, MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            return Const.READ_SUCCESS;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            updatePriority(true);
            updatePriority(false);
        }

        private int updatePriority(bool bAfter)
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
            String insertString;
            if (bAfter)
            {
                insertString = "update clients c set c.priorities_id=12 where c.id in  (select c.id from  clients c where c.dnlimtime0<>'' and " +
                               "c.dnlimtime0<>'00:00' and cast (c.dnlimtime0 as time)<='13:00')";
            }
            else
            {
                insertString = "update clients c set c.priorities_id=13 where c.id in  (select c.id from  clients c where c.dnlimtime0<>'' and " +
                               "c.dnlimtime0<>'00:00' and cast (c.dnlimtime0 as time)>='13:00')";
            }
            FbCommand fbComInsert = new FbCommand(insertString, fbCon);
            fbComInsert.Transaction = fbTrans;
            int insRes = 0;
            try
            {
                insRes = fbComInsert.ExecuteNonQuery();

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message,"Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                fbTrans.Rollback();
                this.Close();
            }
            finally
            {
                fbComInsert.Dispose();
                fbCon.Close();
            }
            return Const.READ_SUCCESS;
        }
    }
}

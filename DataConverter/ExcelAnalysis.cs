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
    public partial class ExcelAnalysis : Form
    {
        String excelOld;
        String excelNew;
        Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetOld;
        Microsoft.Office.Interop.Excel.Application ObjExcelOld;
        Microsoft.Office.Interop.Excel.Workbook ObjWorkBookOld;

        Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheetNew;
        Microsoft.Office.Interop.Excel.Application ObjExcelNew;
        Microsoft.Office.Interop.Excel.Workbook ObjWorkBookNew;
        
        public ExcelAnalysis()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if(radioButton1.Checked == true)
            {
                excelOld = String.Copy(openFileDialog1.FileName);
                label1.Text = excelOld;
            }
            else
            {
                excelNew = String.Copy(openFileDialog1.FileName);
                label2.Text = excelNew;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            List<Data> lDataOld = DownloadDataFromExcel(excelOld, ObjWorkSheetOld,out ObjExcelOld, ObjWorkBookOld);
            List<Data> lDataNew = DownloadDataFromExcel(excelNew, ObjWorkSheetNew,out ObjExcelNew, ObjWorkBookNew);
            List<String> lDataMatch = new List<String>();
            List<Data> lDataMissmatch = new List<Data>();
            //checkingId(lDataOld, lDataNew, lDataMatch);
            //checkingId(lDataNew, lDataOld, lDataMatch);
            int k = 0;
            for (int i = 0; i < lDataNew.Count; i++)
            {
                if (!lDataMatch.Contains(lDataNew[i]._CPCode))
                {
                    k = 0;
                    for (int j = 0; j < lDataOld.Count; j++)
                    {
                        if (!lDataMatch.Contains(lDataOld[j]._CPCode))
                            if (Data.Compare(lDataOld[j], lDataNew[i]) == 1)
                                k++;
                    }
                    if(k==0)
                    {
                        lDataMissmatch.Add(lDataNew[i]);
                    }
                }
            }
            for (int i = 0; i < lDataOld.Count; i++)
            {
                if (!lDataMatch.Contains(lDataOld[i]._CPCode))
                {
                    k = 0;
                    for (int j = 0; j < lDataNew.Count; j++)
                    {
                        if (!lDataMatch.Contains(lDataNew[j]._CPCode))
                            if (Data.Compare(lDataNew[j], lDataOld[i]) == 1)
                                k++;
                    }
                    if (k == 0)
                    {
                        lDataMissmatch.Add(lDataOld[i]);
                    }
                }
            }
            for(int i=0;i<lDataMissmatch.Count;i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1[0, i].Value = lDataMissmatch[i]._CPCode;
                dataGridView1[1, i].Value = lDataMissmatch[i]._CPName;
                dataGridView1[2, i].Value = lDataMissmatch[i]._CPOwner;
                dataGridView1[3, i].Value = lDataMissmatch[i]._CPAdress;
                dataGridView1[4, i].Value = lDataMissmatch[i]._CPSchedule[1];
                dataGridView1[5, i].Value = lDataMissmatch[i]._CPGroupClient;
            }

        }


        List<Data> DownloadDataFromExcel(string excel, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet,out Microsoft.Office.Interop.Excel.Application ObjExcel, Microsoft.Office.Interop.Excel.Workbook ObjWorkBook)
        {
            int i = 10;
            ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            ObjWorkBook = ObjExcel.Workbooks.Open(excel, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            List<Data> lData = new List<Data>();
            Data tmp;
            while (true)
            {
                tmp = new Data();
                i = excelDataToList(i, lData, tmp, ObjWorkSheet);
                if (i != -1)
                {
                    lData.Add(tmp);
                }
                else
                {
                    break;
                }
            }
            return lData;
        }

        private int excelDataToList(int id, List<Data> lData, Data data, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet)
        {
            int iRet = data.dataReadFromExcel(ObjWorkSheet, id, lData);
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
                id = excelDataToList(id, lData, data, ObjWorkSheet);
            }
            return id;
        }

        void checkingId(List<Data> Old, List<Data> New, List<String> Result)
        {
            for (int i = 0; i < Old.Count; i++)
            {
                for (int j = 0; j < New.Count; j++)
                {
                    if (String.Compare(Old[i]._CPCode, New[j]._CPCode) == 0 && !Result.Contains(Old[i]._CPCode))
                    {
                        Result.Add(Old[i]._CPCode);
                        break;
                    }
                }

            }
        }

       
        private void ExcelAnalysis_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                ObjExcelOld.Quit();
                ObjExcelNew.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

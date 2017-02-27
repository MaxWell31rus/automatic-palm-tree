using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DataConverter
{
    [Serializable]
    public class Data
    {
        public String _CPName;
        public String _CPCode;
        public String _CPAdress;
        public List<String> _CPSchedule = new List<string>();
        public String _CPConCount;
        public String _CPType;
        public String _CPOwner;
        public String _CPComment;
        public String _CPStatus;
        public String _CPTimelineB;
        public String _CPTimelineE;
        public String _CPGroupClient;
        public String _Used;
        public String _CPY;
        public String _CPX;
        public int countParams;
        public String CPName
        {
            get { return _CPName; }
        }

        public String CPY
        {
            get { return _CPY; }
        }

        public String CPX
        {
            get { return _CPX; }
        }

        public String CPUsed
        {
            get { return _Used; }
        }
        public String CPCode
        {
            get { return _CPCode; }
        }
        public String CPAdress
        {
            get { return _CPAdress; }
        }
        public List<String> CPSchedule
        {
            get { return _CPSchedule; }
        }
        public String CPConCount
        {
            get { return _CPConCount; }
        }
        public String CPType
        {
            get { return _CPType; }
        }
        public String CPOwner
        {
            get { return _CPOwner; }
        }
        public String CPComment
        {
            get { return _CPComment; }
        }
        public String CPStatus
        {
            get { return _CPStatus; }
        }
        public String CPTimelineB
        {
            get { return _CPTimelineB; }
        }

        public String CPTimelineE
        {
            get { return _CPTimelineE; }
        }

        public String CPGroupClient
        {
            get { return _CPGroupClient; }
        }
        public Data()
        {
            _CPName = "";
            _CPAdress = "";
            _CPCode = "";
            _CPOwner = "";
            _CPConCount = "";
            _CPSchedule.Add("");
            _CPType = "";
            _CPTimelineB = "";
            _CPTimelineE = "";
            _CPStatus = "";
            _CPComment = "";
            _CPGroupClient = "";
            _CPX = "";
            _CPY = "";
            _Used = "1";
            countParams = Const.IDX_ARR.Count();
        }



        public int dataReadFromExcel(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, int i, List<Data> data, List<String> lDuplicate, List<String> lClientsGroups)
        {
            bool flag = true;
            List<Microsoft.Office.Interop.Excel.Range> range = new List<Microsoft.Office.Interop.Excel.Range>();
            try
            {
                for (int j = 0; j < Const.IDX_ARR.Count(); j++)
                {
                    range.Add(ObjWorkSheet.Cells[i, Const.IDX_ARR[j]]);
                }
            }
            catch (Exception e)
            {
                flag = false;
            }
            if (flag)
            {
                if (findId(range[1].Text, data, lDuplicate) == Const.EXIST)
                {
                    return Const.READ_ABORT;
                }
                this._CPCode = range[0].Text;
                this._CPName = range[1].Text;
                this._CPOwner = range[2].Text;
                this._CPSchedule.Add(range[3].Text);
                this.dateConvert();
                this._CPAdress = range[4].Text + ", " + range[5].Text + ", " + range[6].Text;
                this._CPTimelineB = range[7].Text;
                if (_CPTimelineB.Count() > 5)
                {
                    if (this._CPTimelineB[4] != ':')
                    {
                        this._CPTimelineB = this._CPTimelineB.Substring(0, 5);
                    }
                    else
                    {
                        this._CPTimelineB = this._CPTimelineB.Substring(0, 4);
                    }
                }
                this._CPTimelineE = range[8].Text;
                if (_CPTimelineE.Count() > 5)
                {
                    if (this._CPTimelineE[4] != ':')
                    {
                        this._CPTimelineE = this._CPTimelineE.Substring(0, 5);
                    }
                    else
                    {
                        this._CPTimelineE = this._CPTimelineE.Substring(0, 4);
                    }
                }

                if (String.Compare(range[9].Text, "") == 0)
                {

                    this._CPConCount = "1";
                }
                else
                {

                    this._CPConCount = range[9].Text;
                }
                this._CPType = range[10].Text;
                this.typeConvert();
                this._CPComment = range[11].Text;
                if (this._CPComment.Length > 149)
                {
                    string str = this._CPComment.Remove(148);
                    this._CPComment = str;
                }
                this._CPStatus = range[13].Text;
                this.statusConvert();
                if (String.Compare(range[12].Text, "") != 0)
                {
                    this._CPGroupClient = range[12].Text;
                    if (!(lClientsGroups.Contains(range[12].Text)))
                    {
                        lClientsGroups.Add(range[12].Text);
                    }
                }
            }
            else
            {
                this._CPCode = "";
            }
            if (this._CPCode != "")
            {
                return Const.READ_SUCCESS;
            }
            else
            {
                return Const.READ_ERROR;
            }

        }

        public int dataReadFromExcel(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, int i, List<Data> data)
        {
            bool flag = true;
            List<Microsoft.Office.Interop.Excel.Range> range = new List<Microsoft.Office.Interop.Excel.Range>();
            try
            {
                for (int j = 0; j < Const.IDX_ARR.Count(); j++)
                {
                    range.Add(ObjWorkSheet.Cells[i, Const.IDX_ARR[j]]);
                }
            }
            catch (Exception e)
            {
                flag = false;
            }
            if (flag)
            {
                this._CPCode = range[0].Text;
                this._CPName = range[1].Text;
                this._CPOwner = range[2].Text;
                this._CPSchedule.Add(range[3].Text);
                this._CPAdress = range[4].Text + ", " + range[5].Text + ", " + range[6].Text;
                this._CPTimelineB = range[7].Text;
                if (_CPTimelineB.Count() > 5)
                {
                    if (this._CPTimelineB[4] != ':')
                    {
                        this._CPTimelineB = this._CPTimelineB.Substring(0, 5);
                    }
                    else
                    {
                        this._CPTimelineB = this._CPTimelineB.Substring(0, 4);
                    }
                }
                this._CPTimelineE = range[8].Text;
                if (_CPTimelineE.Count() > 5)
                {
                    if (this._CPTimelineE[4] != ':')
                    {
                        this._CPTimelineE = this._CPTimelineE.Substring(0, 5);
                    }
                    else
                    {
                        this._CPTimelineE = this._CPTimelineE.Substring(0, 4);
                    }
                }

                if (String.Compare(range[9].Text, "") == 0)
                {

                    this._CPConCount = "1";
                }
                else
                {

                    this._CPConCount = range[9].Text;
                }
                this._CPType = range[10].Text;
                this.typeConvert();
                this._CPComment = range[11].Text;
                if (this._CPComment.Length > 149)
                {
                    string str = this._CPComment.Remove(148);
                    this._CPComment = str;
                }
                this._CPStatus = range[13].Text;
                this.statusConvert();

                this._CPGroupClient = range[12].Text;


            }
            else
            {
                this._CPCode = "";
            }
            if (this._CPCode != "")
            {
                return Const.READ_SUCCESS;
            }
            else
            {
                return Const.READ_ERROR;
            }

        }


        public void typeConvert()
        {
            char[] str;
            int i = 0;
            if (this._CPType.Count() != 0)
            {
                str = this._CPType.ToCharArray();
            }
            else
            {
                this._CPType = "1";
                return;
            }
            this._CPType = String.Copy("");
            while (i < str.Count())
            {
                if (str[i] != ' ' && str[i] != 'м')
                {
                    if (str[i] == ',')
                    {
                        str[i] = '.';
                    }
                    this._CPType += str[i];
                    i++;
                }
                else
                {
                    i = str.Count();
                }
            }
        }



        static public int findId(String id, List<Data> data, List<String> lDuplicate)
        {
            if (lDuplicate.Contains(id))
            {
                return Const.EXIST;
            }
            for (int i = 0; i < data.Count; i++)
            {
                if (id == data[i].CPCode)
                {
                    lDuplicate.Add(id);
                    data.RemoveAt(i);
                    return Const.EXIST;
                }
            }
            return Const.NOT_EXIST;
        }


        public void dateConvert()
        {
            int flag = 0;
            char[] sep = { ',', ' ' };
            String[] aStr = this.CPSchedule[1].Split(sep);
            _CPSchedule.RemoveAt(1);
            if (aStr.Count() == 1)
            {
                if (Const.DAY_WEEK.Contains(aStr[0]))
                    flag = 2;
            }
            else if (aStr.Count() > 1)
            {
                for (int i = 0; i < aStr.Count(); i++)
                {
                    if (Const.DAY_WEEK.Contains(aStr[i]))
                    {
                        flag = 2;
                    }
                    else
                    {
                        flag = -1;
                        break;
                    }
                }
                if (flag != 2)
                {
                    if (Const.DAY_WEEK.Contains(aStr[1]))
                    {
                        flag = 1;
                    }
                    else
                    {
                        flag = 3;
                    }
                }
            }
            if (flag == 1)
            {
                this._CPSchedule[0] = "1";
                for (int i = 0; i < aStr.Count() - 1; i += 2)
                {
                    this._CPSchedule.Add(Convert.ToInt32(aStr[i]).ToString());
                    this._CPSchedule.Add(aStr[i + 1]);
                }
            }
            else if (flag == 2)
            {
                this._CPSchedule[0] = "2";
                for (int i = 0; i < aStr.Count(); i++)
                {
                    this._CPSchedule.Add(aStr[i]);
                }
            }
            else if (flag == 3)
            {
                this._CPSchedule[0] = "3";
                for (int i = 0; i < aStr.Count() - 2; i++)
                {
                    this._CPSchedule.Add(aStr[i]);
                }
                for (int i = aStr.Count() - 2; i < aStr.Count(); i++)
                {
                    Int32 j = 0;
                    if (Int32.TryParse(aStr[i], out j))
                    {
                        this._CPSchedule.Add(aStr[i]);
                    }
                }
            }

        }

        public static String convertWeekToNmb(String weekDay)
        {
            switch (weekDay)
            {
                case "пн":
                    {
                        return "1";
                    }
                case "вт":
                    {
                        return "2";
                    }
                case "ср":
                    {
                        return "3";
                    }
                case "чт":
                    {
                        return "4";
                    }
                case "пт":
                    {
                        return "5";
                    }
                case "сб":
                    {
                        return "6";
                    }
                case "вс":
                    {
                        return "7";
                    }
                default:
                    {
                        return null;
                    }
            }
        }



        public void statusConvert()
        {
            if (String.Compare(this._CPStatus.ToUpper(), "ПО ГРАФИКУ") == 0)
            {
                this._CPStatus = "1000004";
                this._Used = "1";
            }
            else
            {
                this._CPStatus = "1000003";
                this._Used = "0";
                string[] str = { "3", "1" };
                this._CPSchedule = new List<string>(str);
            }
        }

        public static int Compare(Data Old, Data New)
        {
            if (String.Compare(Old._CPAdress, New._CPAdress) == 0)
                if (String.Compare(Old._CPName, New._CPName) == 0)
                    if (String.Compare(Old._CPOwner, New._CPOwner) == 0)
                        if (String.Compare(Old._CPSchedule[1], New._CPSchedule[1]) == 0)
                            return 1;
                        else
                            return 0;
                    else
                        return 0;
                else
                    return 0;
            else
                return 0;
        }

        ~Data()
        {

        }
    }
}

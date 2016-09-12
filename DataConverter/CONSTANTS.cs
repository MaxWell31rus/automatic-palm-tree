using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConverter
{
    static class Const
    {
        private static int read_success = 1;
        private static int read_abort = 0;
        private static int read_error = -1;
        private static int find_success = 1;
        private static int find_error = -1;
        private static int exist = 1;
        private static int not_exist = -1;
        private static String[] str = { "хмельницкого", "славы", "ватутина", "белгородский", "гражданский", "б.хмельницкого", "б. хмельницкого", "богдана хмельницкого" };
        private static List<String> avenue = new List<String>(str);

        private static String[] str3 = { "1", "2", "3", "4", "5"};
        private static List<String> nmbWeek = new List<String>(str3);

        private static int[] arr= {4 , 6,1, 17, 8, 9, 11, 12, 13, 14, 15, 18, 19, 16};


        private static String[] str1 = { "1", "10а", "10б", "12", "13", "14", "15", "19", "20", "2", "3", "4", "5", "5а", "6", "7", "8", "9", "Белгородский район (правая сторона)", "Великомихайловка", "Строитель гр.1", "Строитель гр.2", "4(новый)", "10а(новый)", "10б(новый)" };
        private static List<String> side = new List<String>(str1);

        private static String[] str2 = { "пн", "вт", "ср", "чт", "пт", "сб", "вс"};
        private static List<String> dayWeek = new List<String>(str2);
        public static int READ_SUCCESS
        {
            get { return read_success; }
        }

        public static int READ_ABORT
        {
            get { return read_abort; }
        }

        public static int READ_ERROR
        {
            get { return read_error; }
        }

        public static int FIND_ERROR
        {
            get { return find_error; }
        }

        public static int FIND_SUCCESS
        {
            get { return find_success; }
        }

        public static int EXIST
        {
            get { return exist; }
        }

        public static int NOT_EXIST
        {
            get { return not_exist; }
        }

        public static List<String> AVENUE
        {
            get {return avenue ;}
        }

        public static List<String> DAY_WEEK
        {
            get { return dayWeek; }
        }

        public static List<String> NMB_WEEK
        {
            get { return nmbWeek; }
        }

        public static List<String> SIDE
        {
            get { return side; }
        }

        public static int[] IDX_ARR
        {
            get { return arr; }
        }

    }
}

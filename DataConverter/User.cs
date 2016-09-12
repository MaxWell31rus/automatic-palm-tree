using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataConverter
{
    class User
    {
        public String login;
        public String password;
        public List<int> rights;

        public User()
        {
            login = null;
            password = null;
            rights = null;
        }

        static public int compare(User usr1,User usr2)
        {
            if (String.Compare(usr1.login, usr2.login) == 0 && String.Compare(usr1.password, usr2.password) == 0)
                return Const.READ_SUCCESS;
            else
                return Const.READ_ERROR;
        }
    }
}

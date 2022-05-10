using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ModelodeAprov.Controller
{
    public class LoginUser
    {
        
        private static String _user;
        private static String _password;



        //get e set

        public static String password

        {

            get { return LoginUser._password; }

            set { LoginUser._password = value; }

        }


        public static String user

        {

            get { return LoginUser._user; }

            set { LoginUser._user = value; }

        }


     
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVDRV
{
    public class Admin
    {
        private static string _name = "Vaness Aerol";
        public static string Name
        {
            get { return _name; }
            set { _name = value; }
        }
    }
}

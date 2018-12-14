using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppk5_v2
{
    class Error_log
    {
        public Error_log(string methodName, Exception e)
        {

        }

        public static string Error_TimeStamp()
        {
            DateTime time = DateTime.Now;
            return string.Format("{0:T}", time);
        }
    }
}

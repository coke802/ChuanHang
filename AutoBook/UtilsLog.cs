using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Diagnostics;

namespace AutoBook
{
    public class UtilsLog
    {
        static public void Log(string Line, params object[] Args)
        {
            string lineText = "AutoBook" + Line;
            string lineLog = String.Format(lineText, Args);
            Debug.Print(lineLog);
        }
    }

}
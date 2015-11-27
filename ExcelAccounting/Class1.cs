using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace ExcelAccounting
{
    public static class MyFunctions
    {
        [ExcelFunction(Description = "Version")]
        public static string XA_Version()
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            return xlApp.Version;
        }

        [ExcelFunction(Description = "Say Hello and Name")]
        public static string XA_SayHello(string name)
        {
            return "Hello " + name + "!";
        }

        [ExcelFunction(Description = "My first .NET function")]
        public static string XA_Price(object[,] array)
        {
            //var caller = Excel(xlfCaller) as ExcelReference;
            //if (caller == null)
            //    return array;

            int rows = array.GetLength(0);
            int columns = array.GetLength(1);
            return String.Format("Hello {0} {1}", rows, columns);
        }
    }
}

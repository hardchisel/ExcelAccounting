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
        public static string XA_Price(object[,] array, string asset_code)
        {
            int rows = array.GetLength(0);
            int cols = array.GetLength(1);

            int x = -1;
            for (int col = 0; col < cols; col++)
            {
                dynamic col_name = array[0, col];
                if (col_name != null && col_name is string && col_name == asset_code)
                {
                    for (int row = 1; row < rows; row++)
                    {
                        dynamic row_name = array[row, 0];
                    }
                    break;
                }
            }

            return String.Format("{0} found at {1}", asset_code, x);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace ExcelAccounting
{
    public struct Asset
    {
        public string code;
        public string name;
        public string base_code;
        public bool invert;
        public bool port;
    }

    public struct Transaction
    {
        public string port;
        public double date;
        public string asset;
        public string account;
        public double amount;
    }

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

        private static Asset GetAsset(object[,] asset_table, string asset_code)
        {
            Asset asset;
            asset.code = "";
            asset.name = "";
            asset.base_code = "";
            asset.invert = false;
            asset.port = false;

            // find asset_code in asset_table
            int rows = asset_table.GetLength(0);
            int cols = asset_table.GetLength(1);

            for (int row = 0; row < rows; row++)
            {
                dynamic code = asset_table[row, 0];
                if (code != null && code is string && code == asset_code)
                {
                    asset.code = asset_table[row, 0].ToString();
                    asset.name = asset_table[row, 1].ToString();
                    asset.base_code = asset_table[row, 2].ToString().ToString();
                    asset.invert = GetBoolean(asset_table[row, 3]);
                    asset.port = GetBoolean(asset_table[row, 4]);
                    break;
                }
            }

            return asset;
        }

        // helper function to convert cell content to boolean
        private static bool GetBoolean(dynamic cell)
        {
            if (cell is double && cell != 0)
                return true;
            else
                return false; 
        }

        //helper function to convert cell content to double
        private static double GetDouble(dynamic cell)
        {
            if (cell is double)
                return cell;
            else
                return 0;
        }

        [ExcelFunction(Description = "Get asset price at specified date")]
        public static double XA_Price(object[,] asset_table, object[,] price_array, string asset_code, double price_date)
        {
            Asset asset = GetAsset(asset_table, asset_code);

            // if asset_code and base are the same, then price = 1 (i.e. USD/USD = 1 always)
            if (asset_code == asset.base_code)
                return 1;

            int price_rows = price_array.GetLength(0);
            int price_cols = price_array.GetLength(1);

            double price = 0;
            for (int price_col = 0; price_col < price_cols; price_col++)
            {
                // check each column header (asset code)
                dynamic col_name = price_array[0, price_col];
                if (col_name != null && col_name is string && col_name == asset_code)
                {
                    // found a matching column, work down rows
                    for (int price_row = 1; price_row < price_rows; price_row++)
                    {
                        // check each row's date
                        dynamic row_date = price_array[price_row, 0];
                        if (row_date != null && row_date is double && row_date <= price_date)
                        {
                            // check if we exceeded requested date
                            if (row_date > price_date)
                                break;
                            else
                            // found a candidate row
                            {
                                if (price_row == (price_rows - 1) && price_date > row_date)
                                {
                                    // we requested a future date
                                    price = 0;
                                    break;
                                }
                                else
                                {
                                    // this could be the price, so save it
                                    dynamic price_candidate = price_array[price_row, price_col];
                                    if (price_candidate != null && price_candidate is double)
                                        price = price_candidate;
                                }
                            }
                        }
                    }
                    break;
                }
            }

            return price;
        }

        [ExcelFunction(Description = "Get asset value at specified date")]
        public static double XA_Value(object[,] asset_table, object[,] price_array, object[,] transaction_table, string asset_code, string value_asset_code, double price_date)
        {
            if (asset_code == value_asset_code)
                return 1;

            Asset asset = GetAsset(asset_table, asset_code);
            if (asset.port)
            {
                // get list of holdings for this collection
                Dictionary<string, double> holdings = new Dictionary<string, double>();
                int transaction_rows = transaction_table.GetLength(0);
                for (int transaction_row = 0; transaction_row < transaction_rows; transaction_row++)
                {
                    Transaction transaction;
                    transaction.port = transaction_table[transaction_row, 0].ToString();
                    transaction.date = GetDouble(transaction_table[transaction_row, 1]);
                    transaction.asset = transaction_table[transaction_row, 2].ToString();
                    transaction.account = transaction_table[transaction_row, 3].ToString();
                    transaction.amount = GetDouble(transaction_table[transaction_row, 4]);

                    if(transaction.port == asset.code && transaction.date <= price_date && transaction.account == "A" )
                    {
                        if (!holdings.ContainsKey(transaction.asset))
                            holdings.Add(transaction.asset, 0);
                        holdings[transaction.asset] += transaction.amount;
                    }
                }

                return holdings.Count;
            }
            else
            {
                double rate = 1;
                // work down from sell asset
                while (asset.base_code != asset.code) // otherwise its just 1
                {
                    double price = XA_Price(asset_table, price_array, asset.code, price_date);
                    if (price == 0)
                        return 0;
                    else if (asset.invert)
                        rate /= price;
                    else
                        rate *= price;
                    if (asset.base_code == value_asset_code)
                        return rate; // complete rate was found
                    asset = GetAsset(asset_table, asset.base_code);
                }
                // work back from value asset to final asset of previous step
                Asset value_asset = GetAsset(asset_table, value_asset_code);
                while (value_asset.base_code != value_asset.code) // otherwise its just 1
                {
                    double price = XA_Price(asset_table, price_array, value_asset.code, price_date);
                    if (price == 0)
                        return 0;
                    else if (value_asset.invert)
                        rate *= price;
                    else
                        rate /= price;
                    if (value_asset.base_code == asset.code)
                        return rate; // complete rate was found
                    value_asset = GetAsset(asset_table, value_asset.base_code);
                }
                // this should never be reached
                return 0;
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace ExcelAccounting
{
    public class Asset
    {
        string code;
        string name;
        string base_code;
        bool invert;
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

        [ExcelFunction(Description = "Get asset price at specified date")]
        public static double XA_Price(object[,] asset_table, object[,] price_array, string asset_code, double price_date)
        {
            // find asset_code in asset_table
            int asset_rows = asset_table.GetLength(0);
            int asset_cols = asset_table.GetLength(1);

            dynamic asset_base = null;
            for (int asset_row = 0; asset_row < asset_rows; asset_row++)
            {
                dynamic code = asset_table[asset_row, 0];
                if (code != null && code is string && code == asset_code)
                { 
                    asset_base = asset_table[asset_row, 2];
                    break;
                }
            }

            // if asset_code and base are the same, then price = 1 (i.e. USD/USD = 1 always)
            if (asset_code == asset_base)
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
        public static double XA_Value(object[,] asset_table, object[,] price_array, string asset_code, string value_asset_code, double price_date)
        {
            if (asset_code == value_asset_code)
                return 1;

            //if (this == valueAsset)
            //    return 1;

            //Asset asset = this;
            //decimal rate = 1;
            //// work down from sell asset
            //while (asset.PricingAsset != asset) // otherwise its just 1
            //{
            //    rate = rate * asset.Price(valueDate);
            //    if (asset.PricingAsset == valueAsset)
            //        return rate; // complete rate was found 
            //    asset = asset.PricingAsset;
            //}
            //// work back from buy asset to final asset of previous step
            //while (valueAsset.PricingAsset != valueAsset) // otherwise its just 1
            //{
            //    rate = rate / valueAsset.Price(valueDate);
            //    if (valueAsset.PricingAsset == asset)
            //        return rate; // complete rate was found 
            //    valueAsset = valueAsset.PricingAsset;
            //}
            //// this should never be reached
            //throw new System.ArgumentException(String.Format("No price found for {0}/{1} at {2}", valueAsset.Code, asset.Code, valueDate.ToShortDateString()));

            return 0;
        }
    }
}

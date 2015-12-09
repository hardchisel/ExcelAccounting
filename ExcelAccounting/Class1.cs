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

        [ExcelFunction(Description = "Get asset by code from asset_table")]
        private static Asset GetAsset(object[,] asset_table, string asset_code)
        {
            Asset asset;
            asset.code = "";
            asset.name = "";
            asset.base_code = "";
            asset.invert = false;

            // find asset_code in asset_table
            int asset_rows = asset_table.GetLength(0);
            int asset_cols = asset_table.GetLength(1);

            for (int asset_row = 0; asset_row < asset_rows; asset_row++)
            {
                dynamic code = asset_table[asset_row, 0];
                if (code != null && code is string && code == asset_code)
                {
                    asset.code = asset_table[asset_row, 0].ToString();
                    asset.name = asset_table[asset_row, 1].ToString();
                    asset.base_code = asset_table[asset_row, 2].ToString();
                    dynamic invert = asset_table[asset_row, 3];
                    if (invert is double && invert != 0)
                        asset.invert = true;
                    break;
                }
            }

            return asset;
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

        [ExcelFunction(Description = "Get asset value at specified date")]
        public static double XA_Value(object[,] asset_table, object[,] price_array, string asset_code, string value_asset_code, double price_date)
        {
            if (asset_code == value_asset_code)
                return 1;

            Asset asset = GetAsset(asset_table, asset_code);
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

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using razorReport.Models;


namespace razorReport.Helper
{
    public class DataHelper{
        public static Dictionary<string, CurrencyType> ConvertData(CashInStockDec stock) {
            var result = new Dictionary<string, CurrencyType>();
            foreach(var prop in stock.GetType().GetProperties()){
                var type = string.Concat(prop.Name.TakeWhile((c) => !Char.IsDigit(c)));
                var key = string.Concat(prop.Name.Reverse().TakeWhile((c) => Char.IsDigit(c)).Reverse());
                if(prop.Name.Contains('-'))
                    key += 'K';
                if(key == "")
                    key = "1K";
                CurrencyType theValue;
                if(!result.TryGetValue(key, out theValue))
                    result[key] = new CurrencyType();
                
                switch (type)
                {
                    case "ax":
                        result[key].Mint = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    case "a":
                        result[key].ATM = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    case "f":
                        result[key].CAD = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    case "u":
                        result[key].CAC = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    case "m":
                        result[key].MUT = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    case "ae":
                        result[key].AE = Convert.ToDouble(prop.GetValue(stock,null));
                        break;
                    default:
                        break;
                }
            }
            return result;
        }
    }
}
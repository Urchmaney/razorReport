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
            var total = new CurrencyType();
            foreach(var prop in stock.GetType().GetProperties()){
                var type = string.Concat(prop.Name.TakeWhile((c) => !Char.IsDigit(c)));
                var key = string.Concat(prop.Name.Reverse().TakeWhile((c) => Char.IsDigit(c)).Reverse());
                if(prop.Name.IndexOf('_') >=0 && key != "")
                    key = (key + "0").Substring(0,2)+"K";
                   
                if(key == "")
                    key = "1K";
                CurrencyType theValue;
                if(!result.TryGetValue(key, out theValue)){
                    theValue = new CurrencyType();
                    result[key] = theValue;
                }
                    
                
                var propVal = Convert.ToDouble(prop.GetValue(stock,null));
                theValue.ATCOB += propVal;
                switch (type)
                {
                    case "ax":
                        theValue.Mint = propVal;
                        total.Mint += propVal;
                        total.ATCOB += propVal;
                        break;
                    case "a":
                        theValue.ATM = propVal;
                        total.ATM += propVal;
                        total.ATCOB += propVal;
                        break;
                    case "f":
                        theValue.CAD = propVal;
                        total.CAD += propVal;
                        total.ATCOB += propVal;
                        break;
                    case "u":
                        theValue.CAC = propVal;
                        total.CAC += propVal;
                        total.ATCOB += propVal;
                        break;
                    case "m":
                        theValue.MUT = propVal;
                        total.MUT   += propVal;
                        total.ATCOB += propVal;
                        break;
                    case "ae":
                        theValue.AE = propVal;
                        total.AE += propVal;
                        total.ATCOB += propVal;
                        break;
                    default:
                        break;
                }
            }
            result.Add("TOTAL", total);
            return result;
        }
    }
}
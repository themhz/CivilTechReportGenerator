using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Helpers {
    public static class MathOperations {

        public static String formatTwoDecimalWithoutRound(String value, int decimal_places = 2) {

            //String a = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString();
            String a = ".";
            try {
                return value.Substring(0, value.IndexOf(a) + decimal_places + 1);
            } catch (Exception e) {
                return value;
            }
            
        }

        
    }
}

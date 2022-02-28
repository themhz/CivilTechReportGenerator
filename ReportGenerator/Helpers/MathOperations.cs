using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Helpers {
    public static class MathOperations {

        public static String formatTwoDecimalWithoutRound(String value, int decimal_places = 2) {
            String decimalPlaces = "";
            for (int i = 0; i < decimal_places; i++)
                decimalPlaces += "0";
            Double v = Double.Parse(value, CultureInfo.InvariantCulture);
            return v.ToString("0."+ decimalPlaces);
        }

        
    }
}

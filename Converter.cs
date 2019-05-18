using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib {
    public static class Converter {
        public static List<T> ConvertDataReaderToList<T>(System.Data.Common.DbDataReader reader) {
            List<T> result = null;



            return result;
        }

        public static int Letter2Num(string columnLetter) {
            if (string.IsNullOrWhiteSpace(columnLetter)) throw new ArgumentNullException("parameter is blank to Letter2Num");

            columnLetter = columnLetter.ToUpperInvariant();

            int result = 0;
            for (int i = 0; i < columnLetter.Length; i++) {
                result *= 26;
                result += columnLetter[i] - 'A' + 1;
            }

            return result;
        }
        public static string Num2Letter(int columnNumber) {
            if (columnNumber <= 0 || columnNumber > 100000) throw new ArgumentOutOfRangeException($"parameter is invalid: {columnNumber}");

            string result = string.Empty;

            int div = columnNumber;
            int mod = 0;

            while (div > 0) {
                mod = (div - 1) % 26;
                result = (char)(65 + mod) + result;
                div = (int)((div - mod) / 26);
            }

            return result;
        }
    }
}

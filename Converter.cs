using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CommonLib {
    public static class Converter {
        //public static List<T> ConvertDataReaderToList<T>(System.Data.Common.DbDataReader reader) {
        //    List<T> result = null;
        //    if (!reader.HasRows) return null;

        //    string colFormatStr = string.Empty;

        //    Dictionary<string, PropertyInfo> col2Prop = new Dictionary<string, PropertyInfo>(StringComparer.InvariantCultureIgnoreCase);
        //    Dictionary<PropertyInfo, string> colFormat = new Dictionary<PropertyInfo, string>();
        //    PropertyInfo[] propertyInfos = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        //    foreach (PropertyInfo propInfo in propertyInfos) {
        //        SqlColumnNameAttribute attr = propInfo.GetCustomAttribute<SqlColumnNameAttribute>();
        //        if (attr != null) {
        //            col2Prop.Add(attr.SqlColumnName, propInfo);

        //            colFormatStr = attr.ColumnFormat;
        //            if (!string.IsNullOrWhiteSpace(colFormatStr)) {
        //                colFormat.Add(propInfo, colFormatStr);
        //            }
        //        }
        //    }

        //    result = new List<T>();
        //    while (reader.Read()) {
        //        T obj = Activator.CreateInstance<T>();
        //        colFormatStr = string.Empty;

        //        for (int i = 0; i < reader.FieldCount; i++) {
        //            string fieldName = reader.GetName(i);
        //            object val = null;
        //            if (col2Prop.ContainsKey(fieldName)) {
        //                PropertyInfo propInfo = col2Prop[fieldName];
        //                object fieldValue = reader.GetValue(i);

        //                if (fieldValue != null) {
        //                    Type fieldType = fieldValue.GetType();
        //                    Type propType = propInfo.PropertyType;

        //                    if (propType != fieldType) {
        //                        if (fieldType == typeof(string) && propType == typeof(DateTime)) {
        //                            colFormatStr = colFormat.ContainsKey(propInfo) ? colFormat[propInfo] : string.Empty;
        //                            val = string.IsNullOrWhiteSpace(colFormatStr) ? DateTime.Parse((string)fieldValue) : DateTime.ParseExact((string)fieldValue, colFormatStr, System.Globalization.CultureInfo.InvariantCulture);
        //                        }
        //                        else {
        //                            val = Convert.ChangeType(fieldValue, propType);
        //                        }
                                 
        //                        propInfo.SetValue(obj, val);
        //                    }
        //                    else {
        //                        propInfo.SetValue(obj, fieldValue);
        //                    }
        //                }
        //                else {
        //                    propInfo.SetValue(obj, null);
        //                }

        //            }
        //            //else if (propertyInfos.Where<PropertyInfo>()) {
        //            else {
        //                PropertyInfo propInfo = typeof(T).GetProperty(fieldName);
        //                if (propInfo != null) {     

        //                    if (fieldValue != null) {
        //                        Type fieldType = fieldValue.GetType();
        //                        Type propType = propInfo.PropertyType;

        //                        if (propType != fieldType) {
        //                            if (fieldType == typeof(string) && propType == typeof(DateTime)) {
        //                                colFormatStr = colFormat.ContainsKey(propInfo) ? colFormat[propInfo] : string.Empty;
        //                                val = string.IsNullOrWhiteSpace(colFormatStr) ? DateTime.Parse((string)fieldValue) : DateTime.ParseExact((string)fieldValue, colFormatStr, System.Globalization.CultureInfo.InvariantCulture);
        //                            }
        //                            else {
        //                                val = Convert.ChangeType(fieldValue, propType);
        //                            }

        //                            propInfo.SetValue(obj, val);
        //                        }
        //                        else {
        //                            propInfo.SetValue(obj, fieldValue);
        //                        }
        //                    }
        //                    else {
        //                        propInfo.SetValue(obj, null);
        //                    }
        //                }
        //            }
        //        }

        //        result.Add(obj);
        //    }


        //    return result;
        //}

        //private static object ConvertReaderValue(object fieldValue, PropertyInfo propInfo) {
        //    object val = null;
        //    //PropertyInfo propInfo = col2Prop[fieldName];
        //    //object fieldValue = reader.GetValue(i);

        //    if (fieldValue != null) {
        //        Type fieldType = fieldValue.GetType();
        //        Type propType = propInfo.PropertyType;

        //        if (propType != fieldType) {
        //            object val = null;
        //            if (fieldType == typeof(string) && propType == typeof(DateTime)) {
        //                colFormatStr = colFormat.ContainsKey(propInfo) ? colFormat[propInfo] : string.Empty;
        //                val = string.IsNullOrWhiteSpace(colFormatStr) ? DateTime.Parse((string)fieldValue) : DateTime.ParseExact((string)fieldValue, colFormatStr, System.Globalization.CultureInfo.InvariantCulture);
        //            }
        //            else {
        //                val = Convert.ChangeType(fieldValue, propType);
        //            }

        //            propInfo.SetValue(obj, val);
        //        }
        //        else {
        //            propInfo.SetValue(obj, fieldValue);
        //        }
        //    }
        //    else {
        //        propInfo.SetValue(obj, null);
        //    }

        //    return val;
        //}

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

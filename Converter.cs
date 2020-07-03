using ADODB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ADODB;
using System.Windows.Forms;

namespace CommonLib
{
    public static class Converter
    {
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

        public static int Letter2Num(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter)) throw new ArgumentNullException("parameter is blank to Letter2Num");

            columnLetter = columnLetter.ToUpperInvariant();

            int result = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                result *= 26;
                result += columnLetter[i] - 'A' + 1;
            }

            return result;
        }
        public static string Num2Letter(int columnNumber)
        {
            if (columnNumber <= 0 || columnNumber > 100000) throw new ArgumentOutOfRangeException($"parameter is invalid: {columnNumber}");

            string result = string.Empty;

            int div = columnNumber;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                result = (char)(65 + mod) + result;
                div = (int)((div - mod) / 26);
            }

            return result;
        }

        internal static List<T> ConvertArrayToList<T>(dynamic[,] arrData, Dictionary<PropertyInfo, int> propColIndex = null, Dictionary<PropertyInfo, string> propColFormat = null)
        {
            List<T> result = null;

            if (arrData == null) return result;
            if (arrData.GetUpperBound(0) < arrData.GetLowerBound(0)) return result;
            if (arrData.Rank != 2) throw new InvalidOperationException("arrData.Rank != 2 ReadData");

            result = new List<T>();
            if (propColIndex != null && propColIndex.Count > 0)
            {
                for (int i = arrData.GetLowerBound(0); i <= arrData.GetUpperBound(0); i++)
                {
                    T obj = Activator.CreateInstance<T>();

                    foreach (KeyValuePair<PropertyInfo, int> item in propColIndex)
                    {
                        PropertyInfo prop = item.Key;
                        int colNum = item.Value;

                        object colValue = arrData[i, colNum];
                        if (colValue != null && prop.PropertyType != colValue.GetType())
                        {
                            object val = null;
                            if (colValue.GetType() == typeof(string) && prop.PropertyType == typeof(DateTime))
                            {
                                string colFormat = propColFormat != null && propColFormat.ContainsKey(prop) ? propColFormat[prop] : string.Empty;
                                val = string.IsNullOrEmpty(colFormat) ? DateTime.Parse((string)colValue) : DateTime.ParseExact((string)colValue, colFormat, System.Globalization.CultureInfo.InvariantCulture);
                            }
                            else
                                val = Convert.ChangeType(colValue, prop.PropertyType);

                            prop.SetValue(obj, val);
                        }
                        else
                            prop.SetValue(obj, colValue);
                    }
                }
            }
            else
            {
                PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
                int maxCol = Math.Min(arrData.GetLength(1), props.Length);
                for (int i = arrData.GetLowerBound(0); i <= arrData.GetUpperBound(0); i++)
                {
                    T obj = Activator.CreateInstance<T>();

                    for (int j = 0; j < maxCol; j++)
                    {
                        object colValue = arrData[i, j + 1];
                        if (colValue != null)
                        {
                            if (colValue.GetType() != props[j].PropertyType)
                            {
                                object val = Convert.ChangeType(colValue, props[j].PropertyType);
                                props[j].SetValue(obj, val);
                            }
                            else
                                props[j].SetValue(obj, colValue);
                        }
                        else
                            props[j].SetValue(obj, null);
                    }
                }
            }
            return result;
        }

        internal static Recordset ConvertListToRecordSet<T>(IEnumerable<T> data)
        {
            Recordset rs = new Recordset();
            rs.CursorLocation = CursorLocationEnum.adUseClient;

            Fields rsFields = rs.Fields;

            PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            HashSet<string> doNotExportProps = new HashSet<string>();

            foreach (PropertyInfo prop in props)
            {
                var att = prop.GetCustomAttribute<DoNotExportToWorksheetAttribute>();
                if (att == null)
                {
                    if (prop.PropertyType == typeof(Guid))
                        rsFields.Append(prop.Name, DataTypeEnum.adVarChar, -1, FieldAttributeEnum.adFldMayBeNull);
                    else
                        rsFields.Append(prop.Name, TranslateToAdodbType(prop.PropertyType), -1, FieldAttributeEnum.adFldMayBeNull);
                }
                else
                    doNotExportProps.Add(prop.Name);
            }

            rs.Open(Missing.Value, Missing.Value, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockOptimistic, 0);
            foreach (T item in data)
            {
                int i = 0;
                rs.AddNew(Missing.Value, Missing.Value);
                for (int j = 0; j < props.Length; j++)
                {
                    PropertyInfo prop = props[j];
                    if (!doNotExportProps.Contains(prop.Name))
                    {
                        if (prop.PropertyType == typeof(Guid))
                            rsFields[i].Value = prop.GetValue(item);
                        else
                        {
                            var val = prop.GetValue(item);
                            if (val != null)
                                rsFields[i].Value = val;
                        }
                        i++;
                    }
                }
            }
            if (rs.RecordCount >0)
                rs.MoveFirst();

            return rs;
        }

        private static DataTypeEnum TranslateToAdodbType(Type propType, int charFieldsMaxLen = 0)
        {
            var t = Nullable.GetUnderlyingType(propType) ?? propType;
            switch (t.UnderlyingSystemType.ToString().ToLower())
            {
                case "system.guid":
                    return DataTypeEnum.adVariant;
                case "system.boolean":
                    return DataTypeEnum.adBoolean;
                case "system.bool":
                    return DataTypeEnum.adBoolean;
                case "system.byte":
                    return DataTypeEnum.adUnsignedTinyInt;
                case "system.char":
                    return DataTypeEnum.adChar;
                case "system.datetime":
                    return DataTypeEnum.adDate;
                case "system.decimal":
                    return DataTypeEnum.adDecimal;
                case "system.double":
                    return DataTypeEnum.adDouble;
                case "system.int16":
                    return DataTypeEnum.adSmallInt;
                case "system.int32":
                    return DataTypeEnum.adInteger;
                case "system.int64":
                    return DataTypeEnum.adBigInt;
                case "system.sbyte":
                    return DataTypeEnum.adTinyInt;
                case "system.single":
                    return DataTypeEnum.adSingle;
                case "system.string":
                    if (charFieldsMaxLen > 4000)
                        return DataTypeEnum.adLongVarWChar;
                    else
                        return DataTypeEnum.adVarWChar;
                case "system.timespan":
                    return DataTypeEnum.adBigInt;
                case "system.uint16":
                    return DataTypeEnum.adUnsignedSmallInt;
                case "system.uint32":
                    return DataTypeEnum.adUnsignedInt;
                case "system.uint64":
                    return DataTypeEnum.adUnsignedBigInt;
                default:
                    break;
            }
            return DataTypeEnum.adVariant;
        }
    }
}


public static class ExcelHelper
{
    /// <summary>
    /// 取得excel 資料
    /// </summary>
    public static IList<T> ExcelToList<T>(string path, string sheetName, int rowStart, string[]? cellVerticalArray)
    {
        List<T> list = new List<T>();
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        FileInfo file = new FileInfo(path);
        using (ExcelPackage excel = new ExcelPackage(file))
        {
            ExcelWorksheet sheet = excel.Workbook.Worksheets[sheetName];

            int rowLast = sheet.Dimension.End.Row + 1;
            for (int i = rowStart; i < rowLast; i++)
            {
                var t = typeof(T);
                T tOject = Activator.CreateInstance<T>();

                var propertieCount = tOject.GetType().GetProperties().Length;
                for (int j = 0; j < propertieCount; j++)
                {
                    var columnNumber = j + 1;
                    var columnName = GetColumnName(columnNumber);
                    if (cellVerticalArray != null && cellVerticalArray.Any())
                    {
                        if (cellVerticalArray.Contains(columnName))
                        {
                            tOject = GetData(sheet, i, tOject, j);
                        }
                    }
                    else
                    {
                        tOject = GetData(sheet, i, tOject, j);
                    }

                }
                list.Add(tOject);
            }
        }

        return list;
    }

    private static T GetData<T>(ExcelWorksheet sheet, int i, T tOject, int j)
    {
        var type = tOject.GetType().GetProperties()[j].PropertyType;
        //var field = GetDescription<T>(tOject.GetType().GetProperties()[j]);

        var cell = (j + 1);
        var val = sheet.Cells[i, cell].Value;

        if (val != null)
        {
            if (sheet.Cells[i, cell].Style.Numberformat.Format.IndexOf("yyyy") > -1
                    && sheet.Cells[i, cell].Value.GetType().ToString() == "System.Double")//處理日期時間格式的關鍵代碼 
            {
                val = sheet.Cells[i, cell].GetValue<DateTime>();
            }


            object r = ChangeType(val, type);
            tOject.GetType().GetProperties()[j].SetValue(tOject, r);
        }


        return tOject;
    }

    /// <summary>
    /// 更新EXCEL 資料
    /// </summary>
    public static void UpdateExcelCellValue<T>(string path, string sheetName, string[]? cellVerticalArray, IList<T> data)
    {
        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        FileInfo file = new FileInfo(path);
        using (ExcelPackage excel = new ExcelPackage(file))
        {
            ExcelWorksheet sheet = excel.Workbook.Worksheets[sheetName];

            var bb = sheet.Dimension.Address;
            int lastRow = sheet.Dimension.End.Row;
            int nowRowIndex = lastRow + 1;

            for (int a = 0; a < data.Count; a++)
            {
                T? item = data[a];
                var properties = (item.GetType().GetProperties());
                for (int i = 0; i < properties.Length; i++)
                {
                    PropertyInfo? propertie = properties[i];
                    //var type = propertie.PropertyType;
                    var val = propertie.GetValue(item);

                    int j = i + 1;
                    if (val != null)
                    {
                        // 日期格式需要做特殊處理
                        if (val.GetType() == typeof(DateTime))
                        {
                            sheet.Cells[nowRowIndex, j].Value = Convert.ToDateTime(val);
                            sheet.Cells[nowRowIndex, j].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        }
                        else
                        {
                            sheet.Cells[nowRowIndex, j].Value = val;
                        }
                    }
                }
                nowRowIndex++;
            }

            // 自動伸縮欄寬
            sheet.Column(2).AutoFit();

            excel.SaveAs(file);
        }
    }

    public static string GetColumnName(int columnNumber)
    {
        var columnLetter = ExcelCellAddress.GetColumnLetter(columnNumber);
        return columnLetter;
    }

    /// <summary>
    /// 寫入資料型別判斷
    /// </summary>
    public static void SetCellValue(ICell cell, object obj)
    {
        if (obj.GetType() == typeof(int))
        {
            cell.SetCellValue((int)obj);
        }
        else if (obj.GetType() == typeof(double))
        {
            cell.SetCellValue((double)obj);
        }
        else if (obj.GetType() == typeof(IRichTextString))
        {
            cell.SetCellValue((IRichTextString)obj);
        }
        else if (obj.GetType() == typeof(string))
        {
            cell.SetCellValue(obj.ToString());
        }
        else if (obj.GetType() == typeof(DateTime))
        {
            cell.SetCellValue((DateTime)obj);
        }
        else if (obj.GetType() == typeof(bool))
        {
            cell.SetCellValue((bool)obj);
        }
        else
        {
            cell.SetCellValue(obj.ToString());
        }
    }

    private static bool CanChangeType(object source, Type targetType)
    {
        try
        {
            Convert.ChangeType(source, targetType);
            return true; // OK, it can be converted
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    private static object ChangeType(object value, Type conversion)
    {
        var t = conversion;

        try
        {
            if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
            {
                if (value == null)
                {
                    return null;
                }

                t = Nullable.GetUnderlyingType(t);
            }

            return Convert.ChangeType(value, t);
        }
        catch (Exception ex)
        {
            return null;
        }
    }

    private static T ChangeType<T>(object value)
    {
        var t = typeof(T);

        if (t.IsGenericType && t.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
        {
            if (value == null)
            {
                return default(T);
            }

            t = Nullable.GetUnderlyingType(t);
        }

        return (T)Convert.ChangeType(value, t);
    }

    private static string GetDescription<T>(PropertyInfo propertyInfo)
    {
        DescriptionAttribute[] attributes = propertyInfo.GetCustomAttributes(typeof(DescriptionAttribute), false) as DescriptionAttribute[];

        if (attributes != null && attributes.Any())
        {
            return attributes.First().Description;
        }

        return string.Empty;
    }
}
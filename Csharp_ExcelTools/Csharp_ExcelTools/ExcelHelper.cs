using NPOI.HSSF.Util;
using System.Diagnostics;

public static class ExcelHelper
{
    /// <summary>
    /// 取得excel 資料
    /// </summary>
    public static IList<T> ExcelToList<T>(string path, string sheetName, int rowStart, string[]? cellVerticalArray)
    {
        List<T> list = new List<T>();
        ISheet sheet;
        using (var stream = new FileStream(path, FileMode.Open))
        {
            stream.Position = 0;
            XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
            sheet = xssWorkbook.GetSheet(sheetName);

            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            //for (int j = 0; j < cellCount; j++)
            //{
            //    ICell cell = headerRow.GetCell(j);
            //    if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
            //    {
            //        // dtTable.Columns.Add(cell.ToString());
            //    }
            //}
            for (int i = (sheet.FirstRowNum + rowStart); i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                var t = typeof(T);
                T tOject = Activator.CreateInstance<T>();
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                    {
                        var cellAddress = row.GetCell(j).Address;
                        if (cellVerticalArray != null && cellVerticalArray.Contains(cellAddress.ToString().Substring(0, 1)))
                        {
                            tOject = GetData(headerRow, row, tOject, j);
                        }
                        else
                        {
                            tOject = GetData(headerRow, row, tOject, j);
                        }
                    }
                }
                list.Add(tOject);
            }
        }
        return list;
    }

    /// <summary>
    /// 更新EXCEL 資料
    /// </summary>
    public static void AddUpdateCellValue<T>(string path, string sheetName, string[]? cellVerticalArray, IList<T> data)
    {
        XSSFWorkbook xssWorkbook;
        ISheet sheet;
        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read))
        {
            stream.Position = 0;
            xssWorkbook = new XSSFWorkbook(stream);
            stream.Close();
        }

        //var stream = new FileStream(path, FileMode.Open, FileAccess.Write);
        //stream.Position = 0;
        //xssWorkbook = new XSSFWorkbook(stream);

        //stream.Position = 0;
        //XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);

        //ICellStyle style1 = xssWorkbook.CreateCellStyle();//樣式
        //style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平對齊方式
        //style1.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直對齊方式
        //                                                                      //設定邊框
        //style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
        //style1.WrapText = true;//自動換行

        //ICellStyle style2 = xssWorkbook.CreateCellStyle();//樣式
        //IFont font1 = xssWorkbook.CreateFont();//字型
        //font1.FontName = "楷體";
        //font1.Color = HSSFColor.Red.Index;//字型顏色
        //font1.Boldweight = (short)FontBoldWeight.Normal;//字型加粗樣式
        //style2.SetFont(font1);//樣式裡的字型設定具體的字型樣式
        //                      //設定背景色
        //style2.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
        //style2.FillPattern = FillPattern.SolidForeground;
        //style2.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
        //style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平對齊方式
        //style2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直對齊方式

        //ICellStyle dateStyle = xssWorkbook.CreateCellStyle();//樣式
        //dateStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;//文字水平對齊方式
        //dateStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;//文字垂直對齊方式
        //                                                                         //設定資料顯示格式
        //IDataFormat dataFormatCustom = xssWorkbook.CreateDataFormat();
        //dateStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

        sheet = xssWorkbook.GetSheet(sheetName);

        ////設定列寬
        //int[] columnWidth = { 10, 10, 20, 10 };
        //for (int i = 0; i < columnWidth.Length; i++)
        //{
        //    //設定列寬度，256*字元數，因為單位是1/256個字元
        //    sheet.SetColumnWidth(i, 256 * columnWidth[i]);
        //}

        IRow headerRow = sheet.GetRow(0);
        int lastRow = sheet.LastRowNum;
        int cellCount = headerRow.LastCellNum;

        //sheet.CreateRow(7).CreateCell(1).SetCellValue("test");

        for (int a = 0; a < data.Count; a++)
        {
            T? item = data[a];
            var properties = (item.GetType().GetProperties());

            int nowRowIndex = lastRow + 1;
            IRow row = sheet.CreateRow(nowRowIndex);
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo? propertie = properties[i];
                var val = propertie.GetValue(item);

                int j = i;

                ICell cell = row.CreateCell(j);
                var cellAddress = row.GetCell(j).Address;
                var columnIndex = row.GetCell(j).ColumnIndex;

                if (val != null)
                {
                    if (cellVerticalArray != null && cellVerticalArray.Contains(cellAddress.ToString().Substring(0, 1)))
                    {
                        cell = sheet.GetRow(nowRowIndex).GetCell(j);
                        SetCellValue(cell, val);
                    }
                    else
                    {
                        cell = sheet.GetRow(nowRowIndex).GetCell(j);
                        SetCellValue(cell, val);
                    }

                    //if (val.GetType() == typeof(DateTime))
                    //{
                    //    cell.CellStyle = dateStyle;
                    //}
                }

            }
            nowRowIndex++;
        }

        using (var stream = new FileStream(path, FileMode.Open, FileAccess.Write))
        {
            //sheet = xssWorkbook.GetSheet(sheetName);
            xssWorkbook.Write(stream);
            stream.Close();
        }
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

    private static T GetData<T>(IRow headerRow, IRow row, T tOject, int j)
    {
        if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
        {
            var type = (tOject.GetType().GetProperties()[j].PropertyType);

            var field = GetDescription<T>(tOject.GetType().GetProperties()[j]);

            if (!string.IsNullOrEmpty(field) && field == headerRow.GetCell(j).ToString())
            {
                var val = row.GetCell(j).ToString();
                if (CanChangeType(val, type))
                {
                    var r = Convert.ChangeType(val, type);
                    tOject.GetType().GetProperties()[j].SetValue(tOject, r);
                }
            }
        }
        return tOject;
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
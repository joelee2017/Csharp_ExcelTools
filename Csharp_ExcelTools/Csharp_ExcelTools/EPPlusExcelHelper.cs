using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace KGI.ExportSHMReport.Helper
{
    public static class EPPlusExcelHelper
    {
        private static string _userName;
        private static string _passWrod;
        private static string _domain;

        /// <summary>
        /// 標楷體
        /// </summary>
        private const string font1 = "標楷體";

        /// <summary>
        /// Calibri
        /// </summary>
        private const string font2 = "Calibri";

        /// <summary>
        /// Times New Roman
        /// </summary>
        private const string font3 = "Times New Roman";

        /// <summary>
        /// 初始化 EPPlusExcelHelper.InitExcelHelper(_logger, _exceptionHelper, _userName, _passWrod, _domain);
        /// </summary>
        public static void InitExcelHelper(
            string userName,
            string passWrod,
            string domain)
        {
            _userName = userName;
            _passWrod = passWrod;
            _domain = domain;
        }

        /// <summary>
        /// 產生Excel EPPlusExcelHelper.ExportExcel(file,data)
        /// </summary>
        private static bool ExportExcel(FileModel model, Employee employee)
        {
            bool Issuccess;
            #region excel 資料寫入
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var source = File.OpenRead(model.TempleateFilePath))
            using (var excel = new ExcelPackage(source))
            {
                ExcelWorksheet ws = excel.Workbook.Worksheets[model.FileName];

                SetContentValue(ws.Cells["A1"], 13, font2, employee.Name);

                SetContentValue(ws.Cells["B1"], 0, font1, employee.Tel);

                SetContentValue(ws.Cells["C1"], 0, font1, employee.Address);

                SetContentValue(ws.Cells["D1"], 0, font1, employee.CreateDate.Value.ToString("yyyy-MM-dd"));

                SetContentValue(ws.Cells["E1"], 0, font1, employee.Number);

                ws.Cells.AutoFitColumns();
                ws.Cells.Style.ShrinkToFit = true;

                excel.Save();

                //儲存Excel
                Issuccess = ExportProcess(model, excel);
            }

            return Issuccess;
            #endregion
        }

        /// <summary>
        /// 寫入值
        /// </summary>
        private static void SetContentValue(ExcelRange cell, int size, string font, object value)
        {
            cell.Value = value;
            cell.Style.Font.Name = font;

            if (size > 0)
                cell.Style.Font.Size = size;
        }

        /// <summary>
        /// 寫入水平置中值
        /// </summary>
        private static void SetContentCenterValue(ExcelRange cell, bool border, int size, object value)
        {
            cell.Value = value;
            cell.Style.Font.Name = font2;
            cell.Style.Font.Size = size;

            if (border)
            {
                cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        /// <summary>
        /// 寫入水平置中值
        /// </summary>
        private static void SetContentRightValue(ExcelRange cell, bool border, int size, object value)
        {
            cell.Value = value;
            cell.Style.Font.Name = font2;
            cell.Style.Font.Size = size;

            if (border)
            {
                cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        }

        /// <summary>
        /// 寫入表頭
        /// </summary>
        private static void SetHeaderValue(ExcelRange cell, int size, string value)
        {
            cell.Value = value;
            cell.Style.Border.Top.Style = ExcelBorderStyle.Medium;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Font.Name = font2;
            cell.Style.Font.Size = size;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid; // 設定背景填色方法
            cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(197, 217, 241));
        }



        /// <summary>
        /// 輸出處理流程
        /// </summary>
        private static bool ExportProcess(FileModel model, ExcelPackage excel)
        {
            byte[] bin = excel.GetAsByteArray();
            string fileName = string.Empty;

            string filePath = model.ExportPath1 + model.ExportPath2 + model.ExportPath3 + model.ExportPath4;

            fileName = model.FileName + ".xlsx";

            bool Issuccess = FileSave(bin, filePath + fileName);

            return Issuccess;
        }

        /// <summary>
        /// 儲存檔案
        /// </summary>
        /// <param name="bin">檔案</param>
        /// <param name="fileName">檔名</param>
        /// <returns></returns>
        private static bool FileSave(byte[] bin, string fileName)
        {
            bool result = default;
            using (Impersonator impersonator = new Impersonator())
            {
                impersonator.Login(_userName, _domain, _passWrod);
                WindowsIdentity.RunImpersonated(impersonator.Identity.AccessToken, () =>
                {
                    File.WriteAllBytes(fileName, bin);
                    result = File.Exists(fileName);
                });
            }

            return result;
        }

        /// <summary>
        /// 創建資料夾  EPPlusExcelHelper.CreateFolder(file)
        /// </summary>
        public static void CreateFolder(FileModel file)
        {
            using (Impersonator impersonator = new Impersonator())
            {
                impersonator.Login(_userName, _domain, _passWrod);
                WindowsIdentity.RunImpersonated(impersonator.Identity.AccessToken, () =>
                {
                    if (!Directory.Exists(file.ExportPath1 + file.ExportPath2))
                    {
                        Directory.CreateDirectory(file.ExportPath1 + file.ExportPath2);

                        Directory.CreateDirectory(file.ExportPath1 + file.ExportPath2 + file.ExportPath3);
                    }

                    // 確認資料夾是否存在
                    string filePath = file.ExportPath1 + file.ExportPath2 + file.ExportPath3;
                    if (Directory.Exists(filePath))
                    {
                        if (!Directory.Exists(filePath))
                        {
                            Directory.CreateDirectory(filePath);
                        }

                        if (Directory.Exists(filePath))
                        {
                            filePath = filePath + file.ExportPath4;
                            if (!Directory.Exists(filePath))
                            {
                                Directory.CreateDirectory(filePath);
                            }

                        }
                    }
                });
            }
        }
    }
}

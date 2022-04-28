
Console.WriteLine("開始讀取並寫入");

// 檔案路徑
string str = @"E:\Project\Csharp_ExcelTools\Csharp_ExcelTools\test1.xlsx";

//包含的 column
string[] cellVerticalArray = { "A" };


// 不指定任何條件即全撈
var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, null);

//var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, cellVerticalArray);


ExcelHelper.UpdateExcelCellValue<Employee>(str, "員工資料", null, result);


Console.WriteLine("結束");
Console.ReadLine();
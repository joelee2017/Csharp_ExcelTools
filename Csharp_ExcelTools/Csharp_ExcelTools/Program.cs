
Console.WriteLine("Hello, World!");

// 檔案路徑
string str = @"E:\Project\Csharp_ExcelTools\Csharp_ExcelTools\test1.xlsx";

//包含的 sheet 以及 cell
string[] cellVerticalArray = { "A" };


// 不指定任何條件即全撈
//var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, null);

var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, cellVerticalArray);


ExcelHelper.AddUpdateCellValue<Employee>(str, "員工資料", null, result);


Console.WriteLine("結束");
Console.ReadLine();
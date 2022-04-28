
Console.WriteLine("開始讀取並寫入");

var config = Appsetting.GetConfigurations();


// 檔案路徑
//string str = @"E:\Project\Csharp_ExcelTools\Csharp_ExcelTools\test1.xlsx";
string str = config.GetRequiredSection("FilePath").Get<string>();

//包含的 column
string[] cellVerticalArray = { "A" };


// 不指定任何條件即全撈
var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, null);

//var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 2, cellVerticalArray);

List<EmployeeExport> employeeExports = new List<EmployeeExport>();
foreach (var item in result)
{
    employeeExports.Add(new EmployeeExport
    {
        Name = item.Name,
        Tel = item.Tel,
        Address = item.Address,
        Number = item.Number,
        CreateDate = item.CreateDate
    });
}

// 不指定任何條件全部寫入
//ExcelHelper.UpdateExcelCellValue<EmployeeExport>(str, "員工資料", null, employeeExports);

ExcelHelper.UpdateExcelCellValue<EmployeeExport>(str, "員工資料", cellVerticalArray, employeeExports);


Console.WriteLine("結束");
Console.ReadLine();
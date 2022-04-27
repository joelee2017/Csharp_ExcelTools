
Console.WriteLine("Hello, World!");

// 檔案路徑
string str = @"E:\Project\Csharp_ExcelTools\Csharp_ExcelTools\test1.xlsx";

//包含的 sheet 以及 cell
string[] cellVerticalArray = { "A", "B", "C", "D", "E" };

//var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料",1, cellVerticalArray);

var result = ExcelHelper.ExcelToList<Employee>(str, "員工資料", 1, null);

// 不指定任何條件即全撈
//var result = ExcelHelper.ReadExcel<Employee>(str, null, null);


//List<EmployeeExport> employees = new List<EmployeeExport>();
//foreach (var item in result)
//{
//    employees.Add(new EmployeeExport {
//        Name = item.Name,
//        Tel = item.Tel,
//        Address = item.Address,
//        Number = item.Number,
//        CreateDate = item.CreateDate
//    });
//}

ExcelHelper.AddUpdateCellValue<Employee>(str, "員工資料", null, result);

//foreach (var item in result)
//{
//    Console.WriteLine($"Name： {item.Name}, Tel： {item.Tel}, Tel： {item.Address}, Number： {item.Number}");
//}

Console.ReadLine();
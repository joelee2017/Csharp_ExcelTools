
Console.WriteLine("開始讀取並寫入");

var config = Appsetting.GetConfigurations();

//// 檔案路徑
//string str = config.GetRequiredSection("FilePath").Get<string>();

//string sheetName = config.GetRequiredSection("SheetName").Get<string>();

////包含的 column
//string[] cellAry = config.GetRequiredSection("CellVertical").Get<string[]>();

//int rowStart = config.GetRequiredSection("RowStart").Get<int>();


List<FileModel> importFiles = config.GetRequiredSection("ImportFiles").Get<List<FileModel>>();

IList<Employee> employees = new List<Employee>();
foreach (var item in importFiles)
{
    employees = ExcelHelper.ExcelToList<Employee>(item.FilePath, item.SheetName, item.RowStart, item.CellVertical);
}


if (employees != null)
{
    List<EmployeeExport> employeeExports = new List<EmployeeExport>();
    foreach (var item in employees)
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

    List<FileModel> exportFiles = config.GetRequiredSection("ExportFiles").Get<List<FileModel>>();
    foreach (var item in exportFiles)
    {
        ExcelHelper.UpdateExcelCellValue<EmployeeExport>(item.FilePath, item.SheetName, item.CellVertical, employeeExports);

    }

    Console.WriteLine("完成結束");
}
else
{
    Console.WriteLine("異常結束，請確認檔案是否存在或超過 5MB");
}

Console.ReadLine();
Environment.Exit(0);
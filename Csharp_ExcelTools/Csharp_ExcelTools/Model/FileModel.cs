public class FileModel
{
    /// <summary>
    /// 檔案名稱
    /// </summary>
    public string FileName { get; set; }

    /// <summary>
    /// 檔案路徑
    /// </summary>
    public string FilePath { get; set; }

    /// <summary>
    /// Sheet 名稱
    /// </summary>
    public string SheetName { get; set; }

    /// <summary>
    /// 開始列
    /// </summary>
    public int RowStart { get; set; }

    /// <summary>
    /// 要包含的欄位
    /// </summary>
    public string[] CellVertical { get; set; }

    /// <summary>
    /// 檔案模版路徑
    /// </summary>
    public string TempleateFilePath { get; set; }

    /// <summary>
    /// 檔案匯出路徑 - 前
    /// </summary>
    public string ExportPath1 { get; set; }

    /// <summary>
    /// 檔案匯出路徑 - 中
    /// </summary>
    public string ExportPath2 { get; set; }


    /// <summary>
    /// 檔案匯出路徑 - 中
    /// </summary>
    public string ExportPath3 { get; set; }


    /// <summary>
    /// 檔案匯出路徑 - 後
    /// </summary>
    public string ExportPath4 { get; set; }
}

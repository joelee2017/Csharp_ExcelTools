public class FileModel
{
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
}
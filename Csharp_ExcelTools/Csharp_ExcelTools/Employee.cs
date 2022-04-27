

public class Employee
{
    [Description("名稱")]
    public string? Name { get; set; }


    [Description("電話")]
    public string? Tel { get; set; }


    [Description("地址")]
    public string? Address { get; set; }

    [Description("數值")]
    public int Number { get; set; }


    [Description("建立日期")]
    public DateTime CreateDate { get; set; }
}

public class EmployeeExport
{
    /// <summary>
    /// 名稱
    /// </summary>
    [Description("A")]
    public string? Name { get; set; }

    /// <summary>
    /// 電話
    /// </summary>

    [Description("B")]
    public string? Tel { get; set; }

    /// <summary>
    /// 地址
    /// </summary>
    [Description("C")]
    public string? Address { get; set; }

    /// <summary>
    /// 數值
    /// </summary>
    [Description("D")]
    public string? Number { get; set; }

    /// <summary>
    /// 建立日期
    /// </summary>
    [Description("E")]
    public string? CreateDate { get; set; }
}
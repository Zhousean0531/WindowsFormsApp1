using System.Collections.Generic;

public class Page5ExportData
{
    // 表單欄位
    public string TestDate { get; set; }
    public string ReportNo { get; set; }
    public string CylinderNo { get; set; }
    public string Customer { get; set; }
    public string FilterType { get; set; }
    public string ReCylinderNo { get; set; }
    public string CarbonLot { get; set; }

    // DGV 解析後的列資料
    public List<Page5RowData> Rows { get; set; }
}

public class Page5RowData
{
    public string SN { get; set; }
    public string Weight { get; set; }

    public Dictionary<int, string> ControlValues { get; set; }

    public string MA { get; set; }
    public string MB { get; set; }
    public string MC { get; set; }
}

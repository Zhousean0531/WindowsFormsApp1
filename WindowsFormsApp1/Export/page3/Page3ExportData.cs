using System.Collections.Generic;

public class Page3ExportData
{
    public string TestingDate { get; set; }
    public string ArrivalDate { get; set; }

    public string CarbonLot { get; set; }
    public string FilterMaterialNo { get; set; }
    public string FilterReportNo { get; set; }
    public string PackageNo { get; set; }
    public string Customer { get; set; }
    public string Model { get; set; }
    public string ReFilterNo { get; set; }
    public string Alarm { get; set; }
    public string WorkOrder { get; set; }
    public string MA { get; set; }
    public string MB { get; set; }
    public string MC { get; set; }

    public string UserName { get; set; }

    public List<Page3RowData> Rows { get; set; } = new List<Page3RowData>();
}
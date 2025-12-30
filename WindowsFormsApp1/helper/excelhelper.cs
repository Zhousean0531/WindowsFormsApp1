public static class ExcelHelper
{
    public static void SetCell(ExcelWorksheet ws, string addr, object value)
    {
        ws.Cells[addr].Value = value ?? "";
    }
}

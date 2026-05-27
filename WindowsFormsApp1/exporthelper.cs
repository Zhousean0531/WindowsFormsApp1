using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;
using WindowsFormsApp1.Data_Access.Page1;
using WindowsFormsApp1.Data_Access.Page2;
using WindowsFormsApp1.Data_Access.Page5;
using WindowsFormsApp1.Data_Access.Page6;
using static Page3ReportExporter;
public static class DiffUtil
{
    public static string GetSumDiff(string out1, string out2, string out3, string in1, string in2, string in3)
    {
        if (double.TryParse(out1, out double o1) &&
            double.TryParse(out2, out double o2) &&
            double.TryParse(out3, out double o3) &&
            double.TryParse(in1, out double i1) &&
            double.TryParse(in2, out double i2) &&
            double.TryParse(in3, out double i3))
        {
            double diff =
                Math.Max(o1 - i1, 0) +
                Math.Max(o2 - i2, 0) +
                Math.Max(o3 - i3, 0);

            return diff > 0 ? diff.ToString("0.###") : "N.D.";
        }

        return "N.D.";
    }
    public static string GetDiff(string outVal, string inVal)
    {
        if (double.TryParse(outVal, out double o) && double.TryParse(inVal, out double i))
        {
            double diff = o - i;
            return diff <= 0 ? "N.D." : diff.ToString("0.###");
        }
        return "N.D.";
    }
}
public static class DictionaryExtensions
{
    public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dict, TKey key, TValue defaultValue = default)
    {
        return dict.TryGetValue(key, out TValue value) ? value : defaultValue;
    }
}
public static class ExportHelper_Page1
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var batch = Page1DataCollector.Collect(tab);
            if (batch == null) return;
            if (!Page1ReportExporter.Export(batch)) return;
            P1Repository.Insert(batch);
            MessageBox.Show("完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page2
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var batch = Page2DataCollector.Collect(tab);
            if (batch == null) return;
            if (!Page2ReportExporter.Export(batch)) return;
            P2Repository.Insert(batch);
            MessageBox.Show("完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page3
{
    public static void Handle(TabPage tab)
    {
        try
        {
            Page3ExportData data = Page3DataCollector.Collect(tab);

            if (data == null)
                return;

            DialogResult result = PressureModeDialog.Show(out Page3PressureMode pressureMode);

            if (result == DialogResult.Cancel)
                return;

            if (!Page3ReportExporter.Export(data, pressureMode))
                return;

            P3Repository.Insert(data);

            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Page3 匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page4
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var batch = Page4DataCollector.Collect(tab);
            if (batch == null) return;
            if (!Page4ReportExporter.Export(batch)) return;
            P4Repository.Insert(batch);
            MessageBox.Show("完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page5
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var data = Page5DataCollector.Collect(tab);
            if (data == null) return;

            string rawEfficiencyText = ControlHelper.GetText(tab, "CYLRawEffTB");

            Page5LookupResult lookupResult = null;

            if (!string.IsNullOrWhiteSpace(data.CylinderNo))
            {
                try
                {
                    lookupResult = Page5LookupHelper.SearchByCylinderNo(data.CylinderNo);
                }
                catch
                {
                    lookupResult = new Page5LookupResult();
                }
            }
            else
            {
                lookupResult = new Page5LookupResult();
            }

            if (!Page5ReportExporter.Export(
                data,
                data.FilterType,
                rawEfficiencyText
            ))
                return;

            P5Repository.Insert(data, rawEfficiencyText);
            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("TabPage5 匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page6
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var data = Page6DataCollector.Collect(tab);
            if (data == null) return;
            if (!Page6ReportExporter.Export(data)) return;
            var batch = Page6DataCollector.CollectForDb(tab, Environment.UserName);
            P6Repository.Insert(batch);
            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}

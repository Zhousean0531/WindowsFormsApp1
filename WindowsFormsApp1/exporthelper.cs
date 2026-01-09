using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Org.BouncyCastle.Pqc.Crypto.Lms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1;
using YourNamespace;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;


public static class DiffUtil
{
    public static string GetSumDiff(string out1, string out2, string out3, string in1, string in2, string in3)
    {
        double o1, o2, o3, i1, i2, i3;
        bool p1 = double.TryParse(out1, out o1);
        bool p2 = double.TryParse(out2, out o2);
        bool p3 = double.TryParse(out3, out o3);
        bool p4 = double.TryParse(in1, out i1);
        bool p5 = double.TryParse(in2, out i2);
        bool p6 = double.TryParse(in3, out i3);

        if (!(p1 && p2 && p3 && p4 && p5 && p6))
            return "N.D.";

        double diff = (o1 + o2 + o3) - (i1 + i2 + i3);
        return diff <= 0 ? "N.D." : diff.ToString("0.###");
    }
    public static string GetDiff(string outVal, string inVal)
    {
        double o, i;
        if (double.TryParse(outVal, out o) && double.TryParse(inVal, out i))
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
            var data = Page1DataCollector.Collect(tab);
            if (data == null) return;

            Page1MasterExporter.Export(data);
            Page1ReportExporter.Export(data);

            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page2
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var data = Page2DataCollector.Collect(tab);

            Page2MasterExporter.Export(data);
            Page2ReportExporter.Export(data);
            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page3
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var data = Page3DataCollector.Collect(tab);
            if (data == null) return;

            Page3MasterExporter.Export(data);

            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}
public static class ExportHelper_Page4
{
    public static void Handle(TabPage tab)
    {
        try
        {
            var data = Page4DataCollector.Collect(tab);
            if (data == null) return;

            Page4MasterExporter.Export(data);
            Page4ReportExporter.Export(data);
            Page4ReportExporterForNanJing.Export(data);
            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
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

            Page5MasterExporter.Export(data);
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
            Page6MasterExporter.Export(data);
            Page6ReportExporter.Export(data);

            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}


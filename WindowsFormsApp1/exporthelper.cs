using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;
using WindowsFormsApp1.Data_Access.Page1;
using WindowsFormsApp1.Data_Access.Page2;
using WindowsFormsApp1.Data_Access.Page4;
using WindowsFormsApp1.Data_Access.Page5;
using WindowsFormsApp1.Data_Access.Page6;
using static Page3ReportExporter;

public class ReportOutputFile
{
    public string TempPath { get; set; }
    public string FinalPath { get; set; }
}

public class ReportExportResult
{
    public bool Success { get; set; }
    public List<ReportOutputFile> Files { get; } = new List<ReportOutputFile>();

    public static ReportExportResult Failed()
    {
        return new ReportExportResult { Success = false };
    }

    public static ReportExportResult FromFile(string tempPath, string finalPath)
    {
        var result = new ReportExportResult { Success = true };
        result.Files.Add(new ReportOutputFile
        {
            TempPath = tempPath,
            FinalPath = finalPath
        });
        return result;
    }
}

public static class ReportExportStaging
{
    public static string CreateTempPath(string finalPath)
    {
        string ext = Path.GetExtension(finalPath);
        if (string.IsNullOrWhiteSpace(ext))
            ext = ".tmp";

        string folder = Path.Combine(
            Path.GetTempPath(),
            "QCReport",
            DateTime.Now.ToString("yyyyMMdd"),
            Guid.NewGuid().ToString("N")
        );

        Directory.CreateDirectory(folder);
        return Path.Combine(folder, "report" + ext);
    }

    public static void Commit(IEnumerable<ReportOutputFile> files)
    {
        foreach (var file in files ?? Enumerable.Empty<ReportOutputFile>())
        {
            if (file == null ||
                string.IsNullOrWhiteSpace(file.TempPath) ||
                string.IsNullOrWhiteSpace(file.FinalPath))
                continue;

            string folder = Path.GetDirectoryName(file.FinalPath);
            if (!string.IsNullOrWhiteSpace(folder))
                Directory.CreateDirectory(folder);

            File.Copy(file.TempPath, file.FinalPath, true);
            TryDelete(file.TempPath);
        }
    }

    public static void Cleanup(IEnumerable<ReportOutputFile> files)
    {
        foreach (var file in files ?? Enumerable.Empty<ReportOutputFile>())
        {
            if (file == null)
                continue;

            TryDelete(file.TempPath);
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                File.Delete(path);
        }
        catch
        {
        }
    }
}

public static class ReportNoHelper
{
    public static bool EnsureRequiredAndAvailable(params string[] reportNos)
    {
        var all = (reportNos ?? new string[0]).Select(x => (x ?? "").Trim()).ToList();

        if (all.Count == 0 || all.Any(string.IsNullOrWhiteSpace))
        {
            MessageBox.Show("有欄位未填入資料，請確認報告編號");
            return false;
        }

        foreach (string reportNo in all.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (Exists(reportNo))
            {
                MessageBox.Show("此報告編號已存在，請確認是否重複執行。");
                return false;
            }
        }

        return true;
    }

    public static void AssignSequentialReportNos(P2Batch batch)
    {
        if (batch == null || batch.GasTests == null || batch.GasTests.Count == 0)
            return;

        for (int i = 0; i < batch.GasTests.Count; i++)
            batch.GasTests[i].ReportNo = Increment(batch.ReportNo, i);
    }

    public static void AssignSequentialReportNos(P4Batch batch)
    {
        if (batch == null || batch.EfficiencyGroups == null || batch.EfficiencyGroups.Count == 0)
            return;

        for (int i = 0; i < batch.EfficiencyGroups.Count; i++)
            batch.EfficiencyGroups[i].ReportNo = Increment(batch.ReportNo, i);
    }

    public static string Increment(string reportNo, int offset)
    {
        reportNo = (reportNo ?? "").Trim();
        if (offset <= 0 || string.IsNullOrWhiteSpace(reportNo))
            return reportNo;

        var match = Regex.Match(reportNo, @"(\d+)$");
        if (!match.Success)
            return reportNo + "-" + (offset + 1).ToString("00");

        string digits = match.Groups[1].Value;
        string prefix = reportNo.Substring(0, reportNo.Length - digits.Length);

        if (!long.TryParse(digits, out long number))
            return reportNo + "-" + (offset + 1).ToString("00");

        return prefix + (number + offset).ToString(new string('0', digits.Length));
    }

    private static bool Exists(string reportNo)
    {
        string sql = @"
DECLARE @ReportNo NVARCHAR(100) = @Value;
DECLARE @Count INT = 0;

SELECT @Count = @Count + COUNT(*) FROM dbo.P1_Batch WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
SELECT @Count = @Count + COUNT(*) FROM dbo.P2_Batch WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
IF COL_LENGTH('dbo.P2_GasTest', 'ReportNo') IS NOT NULL
    SELECT @Count = @Count + COUNT(*) FROM dbo.P2_GasTest WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
SELECT @Count = @Count + COUNT(*) FROM dbo.P3_Batch WHERE LTRIM(RTRIM(ISNULL(FilterReportNo, ''))) = @ReportNo;
SELECT @Count = @Count + COUNT(*) FROM dbo.P4_Batch WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
IF COL_LENGTH('dbo.P4_Efficiency', 'ReportNo') IS NOT NULL
    SELECT @Count = @Count + COUNT(*) FROM dbo.P4_Efficiency WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
SELECT @Count = @Count + COUNT(*) FROM dbo.P5_Batch WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;
SELECT @Count = @Count + COUNT(*) FROM dbo.P6_Batch WHERE LTRIM(RTRIM(ISNULL(ReportNo, ''))) = @ReportNo;

SELECT @Count;";

        using (var conn = DbBootstrap.GetConnection())
        using (var cmd = new SqlCommand(sql, conn))
        {
            cmd.Parameters.AddWithValue("@Value", reportNo);
            conn.Open();
            return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
        }
    }
}
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

            if (!ReportNoHelper.EnsureRequiredAndAvailable(batch.ReportNo))
                return;

            var staged = Page1ReportExporter.ExportStaged(batch);
            if (!staged.Success) return;

            try
            {
                P1Repository.Insert(batch);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

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

            ReportNoHelper.AssignSequentialReportNos(batch);
            if (!ReportNoHelper.EnsureRequiredAndAvailable(
                batch.GasTests.Select(x => x.ReportNo).ToArray()))
                return;

            var staged = Page2ReportExporter.ExportStaged(batch);
            if (!staged.Success) return;

            try
            {
                P2Repository.Insert(batch);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

            var reportNoBox = ControlHelper.Find<TextBox>(tab, "FilterInProcessReportNOTB");
            if (reportNoBox != null)
                reportNoBox.Text = ReportNoHelper.Increment(batch.ReportNo, batch.GasTests.Count);

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

            if (!ReportNoHelper.EnsureRequiredAndAvailable(data.FilterReportNo))
                return;

            var staged = Page3ReportExporter.ExportStaged(data);
            if (!staged.Success)
                return;

            try
            {
                P3Repository.Insert(data);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

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

            ReportNoHelper.AssignSequentialReportNos(batch);
            if (!ReportNoHelper.EnsureRequiredAndAvailable(
                batch.EfficiencyGroups.Select(x => x.ReportNo).ToArray()))
                return;

            var staged = Page4ReportExporter.ExportStaged(batch);
            if (!staged.Success) return;

            try
            {
                P4Repository.Insert(batch);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

            var reportNoBox = ControlHelper.Find<TextBox>(tab, "CylinderRawReportNoTB");
            if (reportNoBox != null)
                reportNoBox.Text = ReportNoHelper.Increment(batch.ReportNo, batch.EfficiencyGroups.Count);

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

            if (!ReportNoHelper.EnsureRequiredAndAvailable(data.ReportNo))
                return;

            var staged = Page5ReportExporter.ExportStaged(
                data,
                data.RawMaterialType,
                rawEfficiencyText
            );

            if (!staged.Success)
                return;

            try
            {
                P5Repository.Insert(data, rawEfficiencyText);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

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

            if (!ReportNoHelper.EnsureRequiredAndAvailable(data.ReportNo))
                return;

            var staged = Page6ReportExporter.ExportStaged(data);
            if (!staged.Success) return;

            var batch = P6Mapper.FromExportData(data, Environment.UserName);

            try
            {
                P6Repository.Insert(batch);
                ReportExportStaging.Commit(staged.Files);
            }
            catch
            {
                ReportExportStaging.Cleanup(staged.Files);
                throw;
            }

            MessageBox.Show("匯出完成");
        }
        catch (Exception ex)
        {
            MessageBox.Show("匯出錯誤：" + ex.Message);
        }
    }
}

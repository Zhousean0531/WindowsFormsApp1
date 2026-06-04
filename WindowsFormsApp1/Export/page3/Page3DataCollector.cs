using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class Page3DataCollector
{
    private const string DgvName = "FilterBox";

    public static Page3ExportData Collect(TabPage tab)
    {
        var data = new Page3ExportData
        {
            TestingDate = ControlHelper.GetText(tab, "FilterTestDateBox"),
            ArrivalDate = ControlHelper.GetText(tab, "FilterProductionBox"),
            WorkOrder = ControlHelper.GetText(tab, "FilterOrderBox"),
            CarbonLot = ControlHelper.GetText(tab, "FilterCarbonLotBox"),
            FilterMaterialNo = ControlHelper.GetText(tab, "FilterMaterialNumerBox"),
            FilterReportNo = ControlHelper.GetText(tab, "FilterReportBox"),
            PackageNo = ControlHelper.GetText(tab, "FilterPackageNOBox"),
            Customer = ControlHelper.GetText(tab, "FilterReportCustmorBox"),
            Model = ControlHelper.GetText(tab, "FilterModelBox"),
            ReFilterNo = ControlHelper.GetText(tab, "ReFilterNOBox"),
            Alarm = ControlHelper.GetText(tab, "FilterAlarmBox"),
            UserName = Environment.UserName
        };
        var eff = P3EfficiencyFinder.FindMinEfficiencyByWorkOrderText(data.CarbonLot);

        data.MA = eff["MA"];
        data.MB = eff["MB"];
        data.MC = eff["MC"];
        var missingFields = new List<string>();

        if (string.IsNullOrWhiteSpace(data.TestingDate))
            missingFields.Add("測試日期");

        if (string.IsNullOrWhiteSpace(data.ArrivalDate))
            missingFields.Add("到貨日期");

        if (string.IsNullOrWhiteSpace(data.FilterReportNo))
            missingFields.Add("報告編號");

        if (string.IsNullOrWhiteSpace(data.Customer))
            missingFields.Add("客戶");

        if (string.IsNullOrWhiteSpace(data.Model))
            missingFields.Add("型號");

        if (missingFields.Any())
        {
            MessageBox.Show(
                "以下基本欄位尚未填寫：\n" + string.Join("、", missingFields),
                "資料未完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return null;
        }

        var dgv = tab.Controls.Find(DgvName, true)
                              .FirstOrDefault() as DataGridView;

        if (dgv == null)
        {
            MessageBox.Show(
                DgvName,
                "系統錯誤",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return null;
        }

        var rows = dgv.Rows.Cast<DataGridViewRow>()
                           .Where(r => !r.IsNewRow)
                           .ToList();

        if (rows.Count == 0)
        {
            MessageBox.Show(
                "尚未輸入任何量測資料。",
                "資料未完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return null;
        }

        if (!ValidateNumericCells(rows))
            return null;

        foreach (var r in rows)
        {
            string sn = GetCell(r, "生產序號");
            string weight = GetCell(r, "重量");

            var rowData = new Page3RowData
            {
                SN = sn,
                Weight = weight,
                Length = GetCell(r, "長"),
                Width = GetCell(r, "寬"),
                Height = GetCell(r, "高"),
                Diagonal = GetCell(r, "對角線"),
                ControlValues = BuildControlValues(r)
            };

            data.Rows.Add(rowData);
        }

        return data;
    }

    private static Dictionary<int, string> BuildControlValues(DataGridViewRow r)
    {
        string particleIn = GetCell(r, "Particle_In", "ParticleIn");
        string particleOut = GetCell(r, "Particle_Out", "ParticleOut");

        string ipaIn = GetCell(r, "IPA_In", "IPAIn");
        string ipaOut = GetCell(r, "IPA_Out", "IPAOut");

        string acetoneIn = GetCell(r, "Acetone_In", "AcetoneIn");
        string acetoneOut = GetCell(r, "Acetone_Out", "AcetoneOut");

        string nonTargetIn = GetCell(r, "Nontarget_In", "NonTargetIn");
        string nonTargetOut = GetCell(r, "Nontarget_Out", "NonTargetOut");

        string pressureDropSpec = GetCell(r, "Pressure_Drop_Spec", "PressureDropSpec");
        string pressureDrop = GetCell(r, "Pressure_Drop", "PressureDrop");

        return new Dictionary<int, string>
        {
            [38] = particleIn,
            [39] = particleOut,
            [40] = DiffUtil.GetDiff(particleOut, particleIn),

            [41] = ipaIn,
            [42] = ipaOut,
            [43] = DiffUtil.GetDiff(ipaOut, ipaIn),

            [44] = acetoneIn,
            [45] = acetoneOut,
            [46] = DiffUtil.GetDiff(acetoneOut, acetoneIn),

            [47] = nonTargetIn,
            [48] = nonTargetOut,
            [49] = DiffUtil.GetDiff(nonTargetOut, nonTargetIn),

            [50] = DiffUtil.GetSumDiff(
                ipaOut, acetoneOut, nonTargetOut,
                ipaIn, acetoneIn, nonTargetIn
            ),

            [53] = pressureDropSpec,
            [54] = pressureDrop
        };
    }

    private static bool ValidateNumericCells(List<DataGridViewRow> rows)
    {
        var fields = new[]
        {
            new NumericField("重量", "重量"),
            new NumericField("長", "長", "length"),
            new NumericField("寬", "寬", "width"),
            new NumericField("高", "高", "height"),
            new NumericField("對角線", "對角線", "diagonal"),
            new NumericField("Particle In", "Particle_In", "ParticleIn"),
            new NumericField("Particle Out", "Particle_Out", "ParticleOut"),
            new NumericField("IPA In", "IPA_In", "IPAIn"),
            new NumericField("IPA Out", "IPA_Out", "IPAOut"),
            new NumericField("Acetone In", "Acetone_In", "AcetoneIn"),
            new NumericField("Acetone Out", "Acetone_Out", "AcetoneOut"),
            new NumericField("Nontarget In", "Nontarget_In", "NonTargetIn"),
            new NumericField("Nontarget Out", "Nontarget_Out", "NonTargetOut"),
            new NumericField("壓損規格", "Pressure_Drop_Spec", "PressureDropSpec"),
            new NumericField("壓損", "Pressure_Drop", "PressureDrop")
        };

        foreach (var row in rows)
        {
            foreach (var field in fields)
            {
                string value = GetCell(row, field.Names).Trim();

                if (string.IsNullOrWhiteSpace(value))
                    continue;

                if (!double.TryParse(value, out _))
                {
                    MessageBox.Show($"{field.DisplayName} 欄位格式錯誤");
                    return false;
                }
            }
        }

        return true;
    }

    private class NumericField
    {
        public NumericField(string displayName, params string[] names)
        {
            DisplayName = displayName;
            Names = names;
        }

        public string DisplayName { get; }
        public string[] Names { get; }
    }

    private static string GetCell(DataGridViewRow row, params string[] names)
    {
        foreach (var name in names)
        {
            foreach (DataGridViewColumn col in row.DataGridView.Columns)
            {
                bool nameMatch = string.Equals(col.Name, name, StringComparison.OrdinalIgnoreCase);
                bool headerMatch = string.Equals(col.HeaderText, name, StringComparison.OrdinalIgnoreCase);

                if (nameMatch || headerMatch)
                {
                    return row.Cells[col.Index]?.Value?.ToString() ?? "";
                }
            }
        }

        return "";
    }
}

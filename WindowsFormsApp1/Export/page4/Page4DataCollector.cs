using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static QuantityHelper;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

public static class Page4DataCollector
{
    public static Page4ExportData Collect(TabPage tab)
    {
        // ─────────────────────────────
        // (A) 基本欄位
        // ─────────────────────────────
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawTestDateBox");
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawArriveDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "CylinderRawTypeBox");
        var reportNoBox = ControlHelper.Find<TextBox>(tab, "CylinderRawReportNoTB");

        if (testPicker == null || arrivePicker == null || materialBox == null || reportNoBox == null)
        {
            MessageBox.Show("缺少必要欄位（日期 / 原料種類 / 報告編號）");
            return null;
        }

        string testingDate = testPicker.Value.ToString("yyyy.MM.dd");
        string arrivalDate = arrivePicker.Value.ToString("yyyy.MM.dd");
        string material = materialBox.Text.Trim();
        string reportNo = reportNoBox.Text.Trim();
        string userName = Environment.UserName;

        var matInfo = MaterialMasterHelper.Get(material);
        string materialNo = matInfo?.MaterialNo ?? "";

        // ─────────────────────────────
        // (B) 多筆資料
        // ─────────────────────────────
        var nos = ParseHelper.SplitStr(ControlHelper.GetText(tab, "CylinderRawNumberBox"));
        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"));

        int n = new[] { nos.Count, weights.Count, vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();
        if (n <= 0)
        {
            MessageBox.Show("沒有可匯出的測試資料");
            return null;
        }

        nos = nos.Take(n).ToList();
        weights = weights.Take(n).ToList();
        vocIns = vocIns.Take(n).ToList();
        vocOuts = vocOuts.Take(n).ToList();
        deltaPs = deltaPs.Take(n).ToList();

        // ─────────────────────────────
        // (C) Lot / 密度 / 數量
        // ─────────────────────────────
        var lotFulls = nos.Select(no =>
        {
            string clean = no.Trim();
            return $"B-{arrivePicker.Value:yyyyMMdd}-001#{clean.PadLeft(2, '0')}";
        }).ToList();

        var lotNos = nos.ToList();

        const double VOL = 50.0;
        var densities = weights.Select(w => w / VOL).ToList();

        string qtyWeight = ControlHelper.GetText(tab, "CylinderRawQtyWeightBox");
        string qtyPack = ControlHelper.GetText(tab, "CylinderRawQtyPackBox");

        string qtyText = QuantityHelper.BuildQuantityText(
            ProductKind.Cylinder,
            material,
            qtyWeight,
            qtyPack
        );

        // ─────────────────────────────
        // (D) 選擇壓損索引
        // ─────────────────────────────
        int selectedIndex;
        using (var f = new Form2(deltaPs, "請選擇本次效率對應的壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;

            selectedIndex = f.SelectedIndex0;
        }

        if (selectedIndex < 0 || selectedIndex >= n)
        {
            MessageBox.Show("選擇的壓損索引不正確");
            return null;
        }

        // ─────────────────────────────
        // (E) Outgassing
        // ─────────────────────────────
        var outgassingList = vocOuts.Zip(vocIns, (o, i) =>
        {
            double diff = o - i;
            return diff <= 0 ? "N.D." : diff.ToString("F1");
        }).ToList();

        // ─────────────────────────────
        // (F) Mesh（只讀 dgv，不再計算）
        // ─────────────────────────────
        var dgv = ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");
        if (dgv == null)
        {
            MessageBox.Show("找不到粒徑資料表");
            return null;
        }

        var particlePercentages = new Dictionary<string, double>();

        foreach (DataGridViewRow row in dgv.Rows)
        {
            if (row.IsNewRow) continue;

            string key = row.Cells[0].Value?.ToString()?.Trim();
            string valText = row.Cells[1].Value?.ToString()?.Trim();

            if (string.IsNullOrWhiteSpace(key)) continue;
            if (key.Contains("總重")) continue;

            if (!double.TryParse(valText, out double percent))
            {
                MessageBox.Show($"粒徑 {key} 的百分比無法解析");
                return null;
            }

            // ⚠️ 這裡「假設第二欄已經是 %」
            particlePercentages[key] = percent;
        }

        string meshSummary = string.Join(" , ",
            particlePercentages.Select(kv => $"{kv.Key} {kv.Value:F1}%")
        );

        // ─────────────────────────────
        // (G) 效率（多氣體 + Panel）
        // ─────────────────────────────
        var effGroups = new List<EfficiencyGroup>();
        var pageType = GasPageType.CylinderRawPage;
        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "CylinderRawEffPanel");

        foreach (var chk in ControlHelper.FindAll<CheckBox>(tab))
        {
            if (!(chk.Tag is string gasKey)) continue;
            if (!chk.Checked) continue;

            var gasCfg = GasMappingHelper.Get(gasKey);
            if (gasCfg == null) continue;
            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui)) continue;

            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (!double.TryParse(tbConc?.Text, out double conc)) continue;
            double.TryParse(tbBg?.Text, out double bg);

            var readings = EfficiencyHelper.ParseReadings(tbVal.Text);
            if (readings.Count < 11) continue;

            var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);
            if (eff == null) continue;

            effGroups.Add(new EfficiencyGroup
            {
                GasName = gasKey,
                Concentration = conc,
                Eff0 = eff.Eff0,
                Eff10 = eff.Eff10,
                Readings11 = eff.Readings,
                Efficiencies11 = eff.Efficiencies
            });
        }

        if (effGroups.Count == 0)
        {
            MessageBox.Show("沒有任何有效的效率資料");
            return null;
        }

        // ─────────────────────────────
        // (H) 回傳 DTO
        // ─────────────────────────────
        return new Page4ExportData
        {
            ReportNo = reportNo,
            Material = material,
            MaterialNo = materialNo,
            ArrivalDate = arrivalDate,
            TestingDate = testingDate,

            LotNos = lotNos,
            LotFulls = lotFulls,
            Densities = densities,
            QtyText = qtyText,
            UserName = userName,

            Weights = weights,
            VocIns = vocIns,
            VocOuts = vocOuts,
            DeltaPs = deltaPs,
            OutgassingList = outgassingList,

            SelectedIndex = selectedIndex,
            ParticleSizePercentages = particlePercentages,
            MeshSummary = meshSummary,
            EfficiencyGroups = effGroups
        };
    }
}

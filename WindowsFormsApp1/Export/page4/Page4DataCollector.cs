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
        if (testPicker == null || arrivePicker == null || materialBox == null)
        {
            MessageBox.Show("找不到測試日期 / 到廠日期 / 原料種類");
            return null;
        }

        string testDate = testPicker.Value.ToString("yyyy.MM.dd");
        string arriveDate = arrivePicker.Value.ToString("yyyy.MM.dd");
        string material = materialBox.Text.Trim();
        

        // 報告編號（允許空）
        string reportNo =
            ControlHelper.Find<TextBox>(tab, "CylinderRawReportNoTB")?.Text.Trim() ?? "";

        // ─────────────────────────────
        // (B) 批號 / 數量
        // ─────────────────────────────
        var lotNos = StringUtil.Split(
            ControlHelper.GetText(tab, "CylinderRawLotBox")
        );

        string qtyWeight = ControlHelper.GetText(tab, "CylinderRawQtyWeightBox");//進貨重量
        string qtyPack = ControlHelper.GetText(tab, "CylinderRawQtyPackBox");//進貨數量
        string qtyText = QuantityHelper.BuildQuantityText(ProductKind.Cylinder, material, qtyWeight, qtyPack);

        // ─────────────────────────────
        // (C) 多筆資料欄位
        // ─────────────────────────────
        var nos = StringUtil.Split(ControlHelper.GetText(tab, "CylinderRawNumberBox"));
        var weights = StringUtil.SplitDouble(ControlHelper.GetText(tab, "CylinderRawWeightBox"));
        var vocIn = StringUtil.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsInletBox"));
        var vocOut = StringUtil.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsOutletBox"));
        var deltas = StringUtil.SplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"));
        var pressure = StringUtil.SplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"));
        var matInfo = MaterialMasterHelper.Get(material);
        string materialNo = matInfo?.MaterialNo ?? "";
        int n = new[] { nos.Count, weights.Count, vocIn.Count, vocOut.Count, deltas.Count }.Min();
        if (n <= 0)
        {
            MessageBox.Show("沒有可匯出的測試資料");
            return null;
        }

        nos = nos.Take(n).ToList();
        weights = weights.Take(n).ToList();
        vocIn = vocIn.Take(n).ToList();
        vocOut = vocOut.Take(n).ToList();
        deltas = deltas.Take(n).ToList();
        lotNos=lotNos.Take(n).ToList();
        pressure = pressure.Take(n).ToList();
        // 密度（固定體積 50）
        const double VOL = 50.0;
        var densities = weights.Select(w => w / VOL).ToList();

        // 完整批號
        var lotFulls = nos.Select(no =>
            $"B-{arrivePicker.Value:yyyyMMdd}-001#{no.PadLeft(2, '0')}"
        ).ToList();

        // Out - In
        var outStrings = vocOut.Zip(vocIn, (o, i) =>
        {
            var d = o - i;
            return d <= 0 ? "N.D." : d.ToString("F1");
        }).ToList();

        // ─────────────────────────────
        // (D) 讓使用者選擇壓損索引
        // ─────────────────────────────
        int selectedIdx;
        using (var f = new Form2(deltas, "請選擇本次效率對應的壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK) return null;
            selectedIdx = f.SelectedIndex0;
        }

        if (selectedIdx < 0 || selectedIdx >= n)
        {
            MessageBox.Show("選擇的壓損索引不合法");
            return null;
        }

        // ─────────────────────────────
        // (E) 粒徑摘要
        // ─────────────────────────────
        string meshSummary = MeshHelper.BuildMeshSummary(tab, material);
        var meshSummaries = Enumerable.Repeat(meshSummary, n).ToList();

        // ─────────────────────────────
        // (F) 效率（多氣體）
        // ─────────────────────────────
        var pageType = GasPageType.CylinderRawPage;
        var effGroups = new List<EfficiencyGroup>();
        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "CylinderRawEffPanel");
        foreach (var chk in ControlHelper.FindAll<System.Windows.Forms.CheckBox>(tab))
        {

            if (!chk.Checked) continue;
            if (!(chk.Tag is string gasKey)) continue;

            var gasCfg = GasMappingHelper.Get(gasKey);
            if (gasCfg == null)
                continue;

            if (!gasCfg.UiMap.TryGetValue(pageType, out var ui))
                continue;

            var uiChk = ControlHelper.Find<System.Windows.Forms.CheckBox>(panel, ui.CheckBoxName); var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);
            if (!double.TryParse(tbConc?.Text, out double conc))
                continue;

            double.TryParse(tbBg?.Text, out double bg);

            var readings = StringUtil.SplitDouble(tbVal?.Text);
            if (readings.Count < 11)
                continue;

            var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);

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
        // (G) 組 DTO
        // ─────────────────────────────
        return new Page4ExportData
        {
            ReportNo = reportNo,
            Material = material,
            ArrivalDate = arriveDate,
            TestingDate = testDate,
            QtyText = qtyText,
            PressureDrops= pressure,
            EfficiencyGroups = effGroups,
            LotFulls = lotFulls,
            MaterialNo = materialNo,
            LotNos =lotNos,
            Densities = densities,
            DeltaPs = deltas,
            VocIns = vocIn,
            Weights=weights,
            VocOuts = vocOut,
            OutgassingList = outStrings,
            SelectedIndex = selectedIdx,
            MeshSummaries = meshSummaries,
        };
    }

    // ─────────────────────────────
    // 小工具：進貨數量字串
    // ─────────────────────────────
    private static string BuildQuantityText(string material, string weightText, string packText)
    {
        bool useLb =
            material.IndexOf("IKP201", StringComparison.OrdinalIgnoreCase) >= 0 ||
            material.IndexOf("IKP205", StringComparison.OrdinalIgnoreCase) >= 0;

        string unit = useLb ? "lb" : "kg";

        string left = string.IsNullOrWhiteSpace(weightText) ? "" : $"{weightText} {unit}";
        string right = string.IsNullOrWhiteSpace(packText) ? "" : $"{packText}包";

        if (left == "" && right == "") return "";
        if (left != "" && right != "") return $"{left}/{right}";
        return left + right;
    }
}

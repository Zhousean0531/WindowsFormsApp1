using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static QuantityHelper;

public static class Page1DataCollector
{
    public static Page1ExportData Collect(TabPage tab)
    {
        // ─────────────────────────────
        // (A) 基本欄位
        // ─────────────────────────────
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawArriveDateBox");
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawTestDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "FilterRawTypeBox");
        var reportNoBox = ControlHelper.Find<TextBox>(tab, "FilterRawReportNoTB");
        var tbConc = ControlHelper.Find<TextBox>(tab, "FilterRawConcertrationBox");
        var tbBg = ControlHelper.Find<TextBox>(tab, "FilterRawBackGroundBox");
        var tbVal = ControlHelper.Find<TextBox>(tab, "FilterRawEffvalueBox");
        var particleSizes = new Dictionary<string, string>();
        
        var dgv = ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox");
        if (dgv != null)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                string sizeName = row.Cells[0].Value?.ToString()?.Trim();
                string sizeValue = row.Cells[1].Value?.ToString()?.Trim();

                if (!string.IsNullOrWhiteSpace(sizeName))
                {
                    particleSizes[sizeName] = sizeValue;
                }
            }
        }

        if (tbConc == null || tbVal == null)
        {
            MessageBox.Show("找不到效率欄位");
            return null;
        }
        if (!double.TryParse(tbConc.Text, out double conc) || conc <= 0)
        {
            MessageBox.Show("初始濃度需 > 0");
            return null;
        }
        double.TryParse(tbBg?.Text, out double bg);

        var readings = EfficiencyHelper.ParseReadings(tbVal.Text);

        if (arrivePicker == null || testPicker == null || materialBox == null)
        {
            MessageBox.Show("缺少必要欄位（日期 / 原料種類）");
            return null;
        }
        string FilterRawQtyWeight =
            ControlHelper.Find<TextBox>(tab, "FilterRawQtyWeightBox")?.Text;
        string FilterRawQuantity =
            ControlHelper.Find<TextBox>(tab, "FilterRawQuantityBox")?.Text;
        string arrivalDate = arrivePicker.Value.ToString("yyyy.MM.dd");
        string testingDate = testPicker.Value.ToString("yyyy.MM.dd");
        string material = materialBox.Text.Trim();
        string reportNo = reportNoBox?.Text.Trim() ?? "";
        string qtyPack = ControlHelper.GetText(tab, "FilterRawQuantityBox");
        string qtyWeight = ControlHelper.GetText(tab, "FilterRawQtyWeight");
        string rawType =
                ControlHelper.Find<TextBox>(tab, "FilterRawTypeBox")?.Text
                ?.Trim()
                ?.ToUpperInvariant();
        string gasName = GasMappingHelper.GetGasNameFromRawType(rawType);
        string qtyText = QuantityHelper.BuildQuantityText(ProductKind.Filter,material, qtyPack, qtyWeight);
        var matInfo = MaterialMasterHelper.Get(material);
        string materialNo = matInfo?.MaterialNo ?? "";
        // ─────────────────────────────
        // (B) 多筆欄位解析
        // ─────────────────────────────
        var nos = ParseHelper.SplitStr(ControlHelper.GetText(tab, "FilterRawNumberBox"));
        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawPressureBox"));
        int n = new[] { nos.Count, weights.Count, vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();
        if (n <= 0)
        {
            MessageBox.Show("沒有可匯出的資料");
            return null;
        }

        nos = nos.Take(n).ToList();
        weights = weights.Take(n).ToList();
        vocIns = vocIns.Take(n).ToList();
        vocOuts = vocOuts.Take(n).ToList();
        deltaPs = deltaPs.Take(n).ToList();

        // ─────────────────────────────
        // (C) 選擇使用哪一筆壓損
        // ─────────────────────────────
        int selectedIndex;
        using (var f = new Form2(deltaPs, "請選擇用哪一筆壓損"))
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
        // (D) 密度 / Outgassing
        // ─────────────────────────────
        const double VOL = 50.0;
        var densities = weights.Select(w => w / VOL).ToList();

        var outgassing = vocOuts.Zip(vocIns, (o, i) =>
        {
            double diff = o - i;
            return diff <= 0 ? "N.D." : diff.ToString("F1");
        }).ToList();

        // ─────────────────────────────
        // (E) 粒徑
        // ─────────────────────────────
        string meshSummary = MeshHelper.BuildMeshSummary(tab, material);
        var meshSummaries = Enumerable.Repeat(meshSummary, n).ToList();

        // ─────────────────────────────
        // (F) 效率
        // ─────────────────────────────
        var effGroups = new List<EfficiencyGroup>();

        var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);

        if (eff != null)
        {
            effGroups.Add(new EfficiencyGroup
            {
                GasName = gasName,
                Result = eff,
                Eff0 = eff.Eff0,
                Eff10 = eff.Eff10,
                Readings11 = eff.Readings,
                Efficiencies11 = eff.Efficiencies
            });
        }

        if (effGroups.Count == 0)
            return null;
        // ─────────────────────────────
        // (G) 組合 Lot 編號
        // ─────────────────────────────
        string arriveKey = arrivePicker.Value.ToString("yyyyMMdd");
        var lotFulls = nos.Select(no =>
            $"B-{arriveKey}-001#{no.PadLeft(2, '0')}"
        ).ToList();

        // ─────────────────────────────
        // (H) 回傳 Page1ExportData
        // ─────────────────────────────
        return new Page1ExportData
        {
            ReportNo = reportNo,
            Material = material,
            ArrivalDate = arrivalDate,
            TestingDate = testingDate,
            FilterRawQtyWeight = FilterRawQtyWeight,
            FilterRawQuantity = FilterRawQuantity,
            MeshSummaries = meshSummaries,
            LotFulls = lotFulls,
            QtyText = qtyText,
            MaterialNo = materialNo,
            Densities = densities,
            DeltaPs = deltaPs,
            VocIns = vocIns,
            VocOuts = vocOuts,
            OutgassingList = outgassing,
            SelectedIndex = selectedIndex,
            Eff0 = eff.Eff0,
            Eff10 = eff.Eff10
        };
    }
}

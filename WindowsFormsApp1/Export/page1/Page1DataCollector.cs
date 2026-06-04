using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page1;
using YourNamespace;
using static QuantityHelper;

public static class Page1DataCollector
{
    public static P1Batch Collect(TabPage tab)
    {
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawArriveDateBox");
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "FilterRawTestDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "FilterRawTypeBox");
        var reportNoBox = ControlHelper.Find<TextBox>(tab, "FilterRawReportNoTB");
        var tbConc = ControlHelper.Find<TextBox>(tab, "FilterRawConcertrationBox")?.Text;
        var tbBg = ControlHelper.Find<TextBox>(tab, "FilterRawBackGroundBox")?.Text;
        var tbVal = ControlHelper.Find<TextBox>(tab, "FilterRawEffvalueBox");
        var batchNo = ControlHelper.Find<TextBox>(tab, "FilterRawBatchNoBox")?.Text;
        var suppliedText = ControlHelper.Find<TextBox>(tab, "FilterRawSuppliedBox")?.Text;
        if (arrivePicker == null || testPicker == null || materialBox == null || reportNoBox == null)
        {
            MessageBox.Show("UI 控制項未找到");
            return null;
        }
        if (tbVal == null)
        {
            MessageBox.Show("找不到效率讀值欄位");
            return null;
        }

        if (!EfficiencyHelper.TryValidateInputs(
            "效率",
            tbConc,
            tbBg,
            tbVal.Text,
            11,
            out double conc,
            out double bg,
            out var readings,
            out string efficiencyError))
        {
            MessageBox.Show(efficiencyError);
            return null;
        }

        var eff = EfficiencyHelper.Compute11Points(conc, bg, readings);
        if (eff == null)
        {
            MessageBox.Show("效率計算失敗");
            return null;
        }

        // ───── 粒徑 ─────
        var dgv = ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox");
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
                MessageBox.Show($"粒徑 {key} 的百分比錯誤");
                return null;
            }

            particlePercentages[key] = percent;
        }

        // ───── 多筆資料 ─────
        var nos = ParseHelper.SplitStr(ControlHelper.GetText(tab, "FilterRawNumberBox"));
        var suppliedList = ParseHelper.SplitStr(suppliedText);

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "FilterRawWeightBox"), out var weights))
        {
            MessageBox.Show("測試品重量欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsInletBox"), out var vocIns))
        {
            MessageBox.Show("Inlet 欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsOutletBox"), out var vocOuts))
        {
            MessageBox.Show("Outlet 欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "FilterRawPressureBox"), out var deltaPs))
        {
            MessageBox.Show("壓損欄位格式錯誤");
            return null;
        }

        string qtyText = QuantityHelper.BuildQuantityText(
            ProductKind.Filter,
            materialBox.Text,
            ControlHelper.GetText(tab, "FilterRawQtyWeight"),
            ControlHelper.GetText(tab, "FilterRawQuantityBox"));

        int n = nos.Count;
        bool countOk =
            n > 0 &&
            weights.Count == n &&
            vocIns.Count == n &&
            vocOuts.Count == n &&
            deltaPs.Count == n;

        if (suppliedList.Count > 1 && suppliedList.Count != n)
            countOk = false;

        if (!countOk)
        {
            MessageBox.Show("各欄位筆數不一致，請確認原料編號、測試品重量、Inlet、Outlet、壓損");
            return null;
        }

        // ───── 選壓損 ─────
        int selectedIndex;
        using (var f = new Form2(deltaPs, "請選擇用哪一筆壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;
            selectedIndex = f.SelectedIndex0;
        }

        var batch = new P1Batch
        {
            ReportNo = reportNoBox.Text.Trim(),
            Material = materialBox.Text.Trim(),
            MaterialNo = MaterialMasterHelper.Get(materialBox.Text)?.MaterialNo ?? "",
            ArrivalDate = arrivePicker.Value,
            TestingDate = testPicker.Value,
            QtyText = qtyText,
            Concentration = (decimal)conc,
            Background = (decimal)bg,
            Username = Environment.UserName,
            ParticleSizePercentages = particlePercentages,
            Samples = new List<P1Sample>() 
        };

        var lotFulls = nos.Take(n)
            .Select(no => $"{batchNo}#{no.PadLeft(2, '0')}")
            .ToList();

        var densities = weights.Take(n).Select(w => w / 50.0).ToList();
        for (int i = 0; i < n; i++)
        {
            decimal? outgassing = null;
            double diff = vocOuts[i] - vocIns[i];
            if (diff > 0)
                outgassing = (decimal)Math.Round(diff, 1);
            var sample = new P1Sample
            {
                LotFull = lotFulls[i],
                SuppliedNO = suppliedList.Count == 1 ? suppliedList[0] : (suppliedList.Count > i ? suppliedList[i] : ""),
                Weight = (decimal?)weights[i],
                Density = (decimal?)densities[i],
                DeltaP = (decimal?)deltaPs[i],
                VocIn = (decimal?)vocIns[i],
                VocOut = (decimal?)vocOuts[i],
                Outgassing = outgassing,
                IsSelected = (i == selectedIndex)
            };

            if (i == selectedIndex)
                sample.Efficiencies = eff.Efficiencies
                    ?.Select(x => (decimal?)x)
                    .ToList();

            batch.Samples.Add(sample);
        }

        return batch;
    }
}

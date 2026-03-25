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
        if (arrivePicker == null || testPicker == null || materialBox == null || reportNoBox == null)
        {
            MessageBox.Show("UI 控制項未找到");
            return null;
        }
        if (!double.TryParse(tbConc, out double conc) || conc <= 0)
        {
            MessageBox.Show("濃度錯誤");
            return null;
        }
        if (!double.TryParse(tbBg, out double bg))
        {
            MessageBox.Show("背景值錯誤");
            return null;
        }
        var readings = EfficiencyHelper.ParseReadings(tbVal.Text);
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
        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "FilterRawPressureBox"));

        string qtyText = QuantityHelper.BuildQuantityText(
            ProductKind.Filter,
            materialBox.Text,
            ControlHelper.GetText(tab, "FilterRawQtyWeight"),
            ControlHelper.GetText(tab, "FilterRawQuantityBox"));

        int n = new[] { nos.Count, weights.Count, vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();
        if (n <= 0)
        {
            MessageBox.Show("資料筆數錯誤");
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

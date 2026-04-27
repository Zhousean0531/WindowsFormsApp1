using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page4;
using WindowsFormsApp1.Helpers;
using YourNamespace;
using static QuantityHelper;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

public static class Page4DataCollector
{
    public static P4Batch Collect(TabPage tab)
    {
        // ===== 基本欄位 =====
        var dtTesting = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawTestDateBox");
        var dtArrival = ControlHelper.Find<DateTimePicker>(tab, "CylinderRawArriveDateBox");
        var txtReportNo = ControlHelper.Find<TextBox>(tab, "CylinderRawReportNoTB");
        var cmbMaterial = ControlHelper.Find<ComboBox>(tab, "CylinderRawTypeBox");

        if (dtTesting == null || dtArrival == null || cmbMaterial == null)
        {
            MessageBox.Show("缺少必要欄位");
            return null;
        }

        string testingDate = dtTesting.Value.ToString("yyyy-MM-dd");
        string arrivalDate = dtArrival.Value.ToString("yyyy-MM-dd");
        string material = cmbMaterial.Text.Trim();
        string reportNo = txtReportNo?.Text.Trim() ?? "";

        string userName = Environment.UserName;

        // ===== Qty =====
        string qtyWeight = ControlHelper.GetText(tab, "CylinderRawQtyWeightBox");
        string qtyPack = ControlHelper.GetText(tab, "CylinderRawQtyPackBox");

        string qtyText = QuantityHelper.BuildQuantityText(
            ProductKind.Cylinder,
            material,
            qtyWeight,
            qtyPack
        );

        // ===== MaterialNo =====
        var matInfo = MaterialMasterHelper.Get(material);
        string materialNo = matInfo?.MaterialNo ?? "";

        // ===== Lot =====
        string supplierLot = ControlHelper.GetText(tab, "CylinderRawLotBox");
        if (string.IsNullOrWhiteSpace(supplierLot))
            supplierLot = "-";

        string factoryLot = ControlHelper.GetText(tab, "CylinderRawLotFullTB");

        // ===== 分割 =====
        var numbers = ParseHelper.SplitStr(ControlHelper.GetText(tab, "CylinderRawNumberBox"));
        var weights = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawWeightBox"));
        var vocIns = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsInletBox"));
        var vocOuts = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsOutletBox"));
        var deltaPs = ParseHelper.SplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"));

        int n = new[] { weights.Count,  vocIns.Count, vocOuts.Count, deltaPs.Count }.Min();

        if (n == 0)
        {
            MessageBox.Show("沒有可用資料");
            return null;
        }

        // ===== Rows =====
        var rows = new List<P4Row>();
        double density = 0;
        
        for (int i = 0; i < n; i++)
        {
            string num = i < numbers.Count ? numbers[i] : "";
            if (i < weights.Count && weights[i] > 0)
            {
                density = Math.Round(weights[i] / 50.0, 3);
            }
            rows.Add(new P4Row
            {
                LotNo = supplierLot,
                LotFull = string.IsNullOrWhiteSpace(num)
                    ? factoryLot
                    : factoryLot + "#" + num,
                Density = density,
                Weight = weights[i],
                VocIn = vocIns[i],
                VocOut = vocOuts[i],
                DeltaP = deltaPs[i],

                Outgassing = (vocOuts[i] - vocIns[i]) <= 0
                    ? "N.D."
                    : (vocOuts[i] - vocIns[i]).ToString("F1"),

                IsSelected = false
            });
        }
        // ===== 壓損選擇 =====
        int selectedIndex;

        using (var f = new Form2(deltaPs, "請選擇壓損"))
        {
            if (f.ShowDialog() != DialogResult.OK)
                return null;

            selectedIndex = f.SelectedIndex0;
        }

        if (selectedIndex >= 0 && selectedIndex < rows.Count)
            rows[selectedIndex].IsSelected = true;

        // ===== Particle =====
        var particlePercentages = new Dictionary<string, double>();

        var dgv = ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");

        if (dgv != null)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;

                string key = row.Cells[0].Value?.ToString();
                string val = row.Cells[1].Value?.ToString();

                if (string.IsNullOrWhiteSpace(key)) continue;

                double percent;
                if (double.TryParse(val, out percent))
                    particlePercentages[key] = percent;
            }
        }

        // ===== Efficiency=====
        var effGroups = new List<P4EfficiencyGroup>();

        var panel = ControlHelper.Find<TableLayoutPanel>(tab, "CylinderRawEffPanel");

        foreach (CheckBox chk in ControlHelper.FindAll<CheckBox>(panel))
        {
            if (!chk.Checked) continue;
            if (!(chk.Tag is string gasKey)) continue;
            var gasCfg = GasMappingHelper.Get(gasKey);
            if (!gasCfg.UiMap.TryGetValue(GasPageType.CylinderRawPage, out var ui)) continue;

            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (tbConc == null || tbVal == null) continue;

            double conc;
            if (!double.TryParse(tbConc.Text, out conc)) continue;

            double bg = 0;
            double.TryParse(tbBg?.Text, out bg);

            var arr = tbVal.Text
                .Replace("／", "/")
                .Replace(",", "/")
                .Replace("\r\n", "/")
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);

            var g = new P4EfficiencyGroup
            {
                GasName = gasKey,
                Concentration = conc
            };

            foreach (var s in arr)
            {
                double v;
                if (double.TryParse(s.Trim(), out v))
                {
                    double eff = 0;

                    if ((conc - bg) != 0)
                        eff = (conc - v) / (conc - bg) * 100.0;

                    g.Efficiencies11.Add(Math.Round(eff, 1));
                }
            }

            if (g.Efficiencies11.Count > 0)
            {
                g.Eff0 = g.Efficiencies11.First();
                g.Eff10 = g.Efficiencies11.Last();
            }

            effGroups.Add(g);
        }

        // ===== Moisture / Butane / Ash =====
        decimal? moisture = null;
        decimal? butane = null;
        decimal? ash = null;

        var chkMoisture = ControlHelper.Find<CheckBox>(tab, "chkMoisture");
        var chkButane = ControlHelper.Find<CheckBox>(tab, "chkButane");
        var chkAsh = ControlHelper.Find<CheckBox>(tab, "chkAsh");

        var tbMoisture = ControlHelper.Find<TextBox>(tab, "CylinderRawMoistureTB");
        var tbButane = ControlHelper.Find<TextBox>(tab, "CylinderRawButaneTB");
        var tbAsh = ControlHelper.Find<TextBox>(tab, "CylinderRawAshTB");

        if (chkMoisture != null && chkMoisture.Checked)
        {
            decimal v;
            if (decimal.TryParse(tbMoisture?.Text, out v))
                moisture = v;
        }

        if (chkButane != null && chkButane.Checked)
        {
            decimal v;
            if (decimal.TryParse(tbButane?.Text, out v))
                butane = v;
        }

        if (chkAsh != null && chkAsh.Checked)
        {
            decimal v;
            if (decimal.TryParse(tbAsh?.Text, out v))
                ash = v;
        }
        // ===== 回傳 =====
        return new P4Batch
        {
            ReportNo = reportNo,
            Material = material,
            MaterialNo = materialNo,
            ArrivalDate = arrivalDate,
            TestingDate = testingDate,
            UserName = userName,
            QtyText = qtyText,

            Rows = rows,
            ParticleSizePercentages = particlePercentages,
            EfficiencyGroups = effGroups,
            Density = density,
            Moisture = moisture,
            Butane = butane,
            Ash = ash
        };
    }
}
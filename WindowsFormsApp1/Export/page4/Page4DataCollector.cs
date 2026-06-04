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

        string factoryLot = ControlHelper.GetText(tab, "CylinderRawLotFullTB");

        // ===== 分割 =====
        var numbers = ParseHelper.SplitStr(ControlHelper.GetText(tab, "CylinderRawNumberBox"));
        var supplierLots = ParseHelper.SplitStr(supplierLot);

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "CylinderRawWeightBox"), out var weights))
        {
            MessageBox.Show("測試品重量欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsInletBox"), out var vocIns))
        {
            MessageBox.Show("Inlet 欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "CylinderRawVOCsOutletBox"), out var vocOuts))
        {
            MessageBox.Show("Outlet 欄位格式錯誤");
            return null;
        }

        if (!ParseHelper.TrySplitDouble(ControlHelper.GetText(tab, "CylinderRawPressureBox"), out var deltaPs))
        {
            MessageBox.Show("壓損欄位格式錯誤");
            return null;
        }

        int n = numbers.Count;
        bool countOk =
            n > 0 &&
            weights.Count == n &&
            vocIns.Count == n &&
            vocOuts.Count == n &&
            deltaPs.Count == n;

        if (supplierLots.Count > 1 && supplierLots.Count != n)
            countOk = false;

        if (!countOk)
        {
            MessageBox.Show("各欄位筆數不一致，請確認原料編號、測試品重量、Inlet、Outlet、壓損");
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
                LotNo = supplierLots.Count == 1 ? supplierLots[0] : (supplierLots.Count > i ? supplierLots[i] : ""),
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

                string key = row.Cells[0].Value?.ToString()?.Trim();
                string val = row.Cells[1].Value?.ToString()?.Trim();

                if (string.IsNullOrWhiteSpace(key)) continue;
                if (key.Contains("總重")) continue;

                double percent;
                if (double.TryParse(val, out percent))
                    particlePercentages[key] = percent;
                else
                {
                    MessageBox.Show($"粒徑 {key} 的百分比錯誤");
                    return null;
                }
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
            if (gasCfg == null) continue;
            if (!gasCfg.UiMap.TryGetValue(GasPageType.CylinderRawPage, out var ui)) continue;

            var tbConc = ControlHelper.Find<TextBox>(panel, ui.ConcBox);
            var tbBg = ControlHelper.Find<TextBox>(panel, ui.BgBox);
            var tbVal = ControlHelper.Find<TextBox>(panel, ui.ValueBox);

            if (tbConc == null || tbBg == null || tbVal == null)
            {
                MessageBox.Show($"{gasKey} 找不到濃度 / 背景 / 讀值欄位");
                return null;
            }

            if (!EfficiencyHelper.TryValidateInputs(
                gasKey,
                tbConc.Text,
                tbBg.Text,
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

            var g = new P4EfficiencyGroup
            {
                GasName = gasKey,
                Concentration = conc
            };

            foreach (var reading in readings)
            {
                double v = reading.Value;
                double eff = 0;

                if ((conc - bg) != 0)
                    eff = (conc - v) / (conc - bg) * 100.0;

                g.Efficiencies11.Add(Math.Round(eff, 1));
            }

            if (g.Efficiencies11.Count > 0)
            {
                g.Eff0 = g.Efficiencies11.First();
                g.Eff10 = g.Efficiencies11.Last();
            }

            effGroups.Add(g);
        }

        // ===== Moisture / Butane / Ash =====
        string moisture = null;
        string butane = null;
        string ash = null;

        var chkMoisture = ControlHelper.Find<CheckBox>(tab, "chkMoisture");
        var chkButane = ControlHelper.Find<CheckBox>(tab, "chkButane");
        var chkAsh = ControlHelper.Find<CheckBox>(tab, "chkAsh");

        var tbMoisture = ControlHelper.Find<TextBox>(tab, "CylinderRawMoistureTB");
        var tbButane = ControlHelper.Find<TextBox>(tab, "CylinderRawButaneTB");
        var tbAsh = ControlHelper.Find<TextBox>(tab, "CylinderRawAshTB");

        if (chkMoisture != null && chkMoisture.Checked)
        {
            string input = (tbMoisture?.Text ?? "").Trim();

            if (input == "-")
                moisture = input;
            else if (string.IsNullOrWhiteSpace(input))
                moisture = null;
            else
            {
                decimal v;
                if (decimal.TryParse(input, out v))
                    moisture = v.ToString("0.###");
                else
                {
                    MessageBox.Show("水分只能輸入數字或 -");
                    return null;
                }
            }
        }

        if (chkButane != null && chkButane.Checked)
        {
            string input = (tbButane?.Text ?? "").Trim();

            if (input == "-")
                butane = input;
            else if (string.IsNullOrWhiteSpace(input))
                butane = null;
            else
            {
                decimal v;
                if (decimal.TryParse(input, out v))
                    butane = v.ToString("0.###");
                else
                {
                    MessageBox.Show("正丁烷只能輸入數字或 -");
                    return null;
                }
            }
        }

            if (chkAsh != null && chkAsh.Checked)
        {
            string input = (tbAsh?.Text ?? "").Trim();

            if (input == "-")
                ash = input;
            else if (string.IsNullOrWhiteSpace(input))
                ash = null;
            else
            {
                decimal v;
                if (decimal.TryParse(input, out v))
                    ash = v.ToString("0.###");
                else
                {
                    MessageBox.Show("灰分只能輸入數字或 -");
                    return null;
                }
            }
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

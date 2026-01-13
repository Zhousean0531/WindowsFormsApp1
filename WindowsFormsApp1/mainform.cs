using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        private Page5LookupResult _page5LookupResult;
        private const int MAX_ROWS = 13;
        public Form1()
        {
            InitializeComponent();
            this.FilterRawTypeBox.SelectedIndexChanged += new System.EventHandler(this.FilterRawTypeBox_SelectedIndexChanged);
            string today = DateTime.Now.ToString("yyyy.MM.dd");
            CylinderRawMoistureTB.Click += TxtMoisture_Click;
            CylinderRawAshTB.Click += TxtAsh_Click;
            // 勾勾變化時，切換 TextBox 可輸入狀態
            chkMoisture.CheckedChanged += (s, e) => ToggleMoistureUI();
            chkAsh.CheckedChanged += (s, e) => ToggleAshUI();
            ToggleMoistureUI();
            ToggleAshUI();
        }
        private void ClearPage5()
        {
            CylinderBox.Rows.Clear();
            CylinderReportNOBox.Clear();
            CylinderCustmorBox.Clear();
            ReCylinderBox.Clear();
            CYLTypeBox.Text = "";
            CYLRawMaterialBox.Clear();
            CYLRawEffTB.Clear();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            MaterialMasterHelper.Load();
        }
        private void FilterRawTypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string m = (FilterRawTypeBox.Text ?? "").Trim().ToUpperInvariant();

            FilterRawParticleSizeBox.Rows.Clear();
            FilterRawParticleSizeBox.AllowUserToAddRows = false;

            if (m.Contains("SG") || m == "IKP101")
            {
                FilterRawParticleSizeBox.Rows.Add("總重");
                FilterRawParticleSizeBox.Rows.Add("<30 Mesh");
                FilterRawParticleSizeBox.Rows.Add("30~60 Mesh");
                FilterRawParticleSizeBox.Rows.Add(">60 Mesh");

            }
            else if (m == "SI013")
            {
                FilterRawParticleSizeBox.Rows.Add("總重");
                FilterRawParticleSizeBox.Rows.Add("<30 Mesh");
                FilterRawParticleSizeBox.Rows.Add("30~70 Mesh");
                FilterRawParticleSizeBox.Rows.Add(">70 Mesh");
            }
        }
        private void FilterRawTestDateBox_ValueChange(object sender, EventArgs e)
        {
            var FilterRawReportNo = FilterRawTestDateBox.Value.ToString("yyMMdd");
            FilterRawReportNoTB.Text = $"YS-Q-M11-" + FilterRawReportNo;
        }
        private void CylinderRawTestDateBox_ValueChange(object sender, EventArgs e)
        {
            var CylinderRawReportNo = CylinderRawTestDateBox.Value.ToString("yyMMdd");
            CylinderRawReportNoTB.Text = $"YS-Q-C11-" + CylinderRawReportNo;
        }
        private void FilterInProcessTestBox_ValueChange(object sender, EventArgs e)
        {
            var FilterSemiRawReportNo = FilterInProcessTestBox.Value.ToString("yyMMdd");
            FilterInProcessReportNOTB.Text = $"YS-Q-M20-" + FilterSemiRawReportNo;
        }
        private void FillDgvFromLookup(Page5LookupResult result)
        {
            var dgv = CylinderBox;

            dgv.Rows.Clear();

            foreach (var rowData in result.Rows)
            {
                int rowIndex = dgv.Rows.Add();
                var dgvRow = dgv.Rows[rowIndex];

                foreach (var kv in rowData)
                {
                    dgvRow.Cells[kv.Key - 1].Value = kv.Value;
                }
            }
        }
        private void FilterRawArriveDateBox_ValueChanged(object sender, EventArgs e)
        {
            var dt = FilterRawArriveDateBox.Value.ToString("yyyyMMdd");
            FilterRawBatchNOBox.Text = $"B-{dt}-001";
        }
        private void FilterProductTestDateBox_ValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = FilterTestDateBox.Value;

            // 格式：YS-Q-M30-yymmdd（yy 為西元後兩碼）
            string formattedDate = selectedDate.ToString("yyMMdd");

            FilterReportBox.Text = "YS-Q-M30-" + formattedDate;
        }
        private void CylinderTestDateBox_ValueChanged(object sender, EventArgs e)
        {
            DateTime selectedDate = CylinderTestDateBox.Value;

            // 格式：YS-Q-M30-yymmdd（yy 為西元後兩碼）
            string formattedDate = selectedDate.ToString("yyMMdd");

            CylinderReportNOBox.Text = "YS-Q-C30-" + formattedDate;
        }
        private void MaterialTestDateBox_ValueChanged(object sender, EventArgs e)
        {
            if (MaterialTypeTB.SelectedItem == null) return;

            string typeCode = MaterialTypeTB.Text == "濾網" ? "M" : "C";
            string datePart = MaterialTestDateBox.Value.ToString("yyMMdd");

            MaterialReportNOTB.Text = $"YS-Q-{typeCode}12-{datePart}";
        }
        private void CylinderBox_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            // 計算行號（1-based）
            string rowNumber = (e.RowIndex + 1).ToString();

            // 定位行號繪製位置
            var centerFormat = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            Rectangle headerBounds = new Rectangle(
                e.RowBounds.Left,
                e.RowBounds.Top,
                FilterBox.RowHeadersWidth,
                e.RowBounds.Height);

            // 寫上數字到左邊 row header
            e.Graphics.DrawString(rowNumber, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
        private void CylinderRawTypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string a = (CylinderRawTypeBox.Text ?? "").Trim().ToUpperInvariant();

            CylinderRawMeshBox.Rows.Clear();
            CylinderRawMeshBox.AllowUserToAddRows = false;

            if (a.Contains("IKP"))
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">7.1mm");
                CylinderRawMeshBox.Rows.Add("5.6~7.1mm");
                CylinderRawMeshBox.Rows.Add("2.8~5.6mm");
                CylinderRawMeshBox.Rows.Add("2.36~2.8mm");
                CylinderRawMeshBox.Rows.Add("<2.36mm");
            }
            else if (a == "CI001")
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">5mm");
                CylinderRawMeshBox.Rows.Add("3.35~5mm");
                CylinderRawMeshBox.Rows.Add("2.36~3.35mm");
                CylinderRawMeshBox.Rows.Add("<2.36mm");
            }
            else if (a == "SG017_A" || a == "SG007" || a == "SG043" || a == "SG029")
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">4.75mm");
                CylinderRawMeshBox.Rows.Add("2.36~4.75mm");
                CylinderRawMeshBox.Rows.Add("<2.36mm");
            }
            else if (a == "SG017_B")
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">3.35mm");
                CylinderRawMeshBox.Rows.Add("2~3.35mm");
                CylinderRawMeshBox.Rows.Add("<2mm");
            }
            else if (a == "SG017_C")
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">3.35mm");
                CylinderRawMeshBox.Rows.Add("2~3.35mm");
                CylinderRawMeshBox.Rows.Add("1.7~2.mm");
                CylinderRawMeshBox.Rows.Add("<1.7mm");
            }
            else if (a == "SG017_D" || a == "SG027")
            {
                CylinderRawMeshBox.Rows.Add("總重");
                CylinderRawMeshBox.Rows.Add(">4.75mm");
                CylinderRawMeshBox.Rows.Add("1.7~4.75mm");
                CylinderRawMeshBox.Rows.Add("<1.7mm");
            }
        }
        private void Execute_Click(object sender, EventArgs e)
        {
            var currentTab = tabControl1.SelectedTab;
            // --- 校正檢查區 ---
            List<int> columnsToCheck = null;

            switch (currentTab.Name)
            {
                case "FilterPage":   // 濾網成品
                    columnsToCheck = new List<int> { 1, 2, 3, 4, 5, 6 }; // A~F
                    break;

                case "CylinderPage": // 濾筒成品
                    columnsToCheck = new List<int> { 1, 2, 3, 6, 7, 8 }; // A~C, F, G, H
                    break;
            }
            // 如果需要校正檢查 → 執行檢查
            if (columnsToCheck != null)
            {
                if (!CalibrationHelper.CheckCalibrationByColumns(columnsToCheck))
                    return; // 校正不合格 → 停止匯出
            }
            switch (currentTab.Name)
            {
                case "FilterRawPage":
                    ExportHelper_Page1.Handle(currentTab);
                    break;
                case "FilterInProcessPage":
                    ExportHelper_Page2.Handle(currentTab);
                    break;
                case "FilterPage":
                    ExportHelper_Page3.Handle(currentTab);
                    break;
                case "CylinderRawPage":
                    ExportHelper_Page4.Handle(currentTab);
                    break;
                case "CylinderPage":
                    ExportHelper_Page5.Handle(currentTab);
                    break;
                case "RawMaterialPage":
                    ExportHelper_Page6.Handle(currentTab);
                    break;
                default:
                    MessageBox.Show("找不到對應頁面模組");
                    break;
            }
            var sb = new StringBuilder();
            sb.AppendLine($"[目前分頁]: {currentTab.Name}");

            foreach (System.Windows.Forms.Control control in currentTab.Controls)
            {
                if (control is TextBox tb)
                {
                    sb.AppendLine($"{tb.Name} = {tb.Text}");
                }
                else if (control is ComboBox cb)
                {
                    sb.AppendLine($"{cb.Name} = {cb.SelectedItem}");
                }
                else if (control is DateTimePicker dt)
                {
                    sb.AppendLine($"{dt.Name} = {dt.Value:yyyy.MM.dd}");
                }
                else if (control is DataGridView dgv)
                {
                    sb.AppendLine($"[{dgv.Name} 資料表]");
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        if (!row.IsNewRow)
                        {
                            List<string> rowData = new List<string>();
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                rowData.Add(cell.Value?.ToString() ?? "");
                            }
                            sb.AppendLine(string.Join(" | ", rowData));
                        }
                    }
                }
            }
        }
        private void ToggleMoistureUI()
        {
            bool on = chkMoisture.Checked;
            if (!on) CylinderRawMoistureTB.Clear();
        }
        private void ToggleAshUI()
        {
            bool on = chkAsh.Checked;
            if (!on) CylinderRawAshTB.Clear();
        }
        private void TxtMoisture_Click(object sender, EventArgs e)
        {
            // 沒勾選 → 當一般 TextBox，不做任何事
            if (!chkMoisture.Checked)
                return;

            // 有勾選 → 攔截點擊，改成開計算視窗
            using (var f = new FormCalcMoistureAsh(CalcMode.Moisture))
            {
                if (f.ShowDialog(this) == DialogResult.OK &&
                    f.AverageResult.HasValue)
                {
                    CylinderRawMoistureTB.Text =
                        f.AverageResult.Value.ToString("F2");
                }
            }
        }
        private void TxtAsh_Click(object sender, EventArgs e)
        {
            if (!chkAsh.Checked)
                return;

            using (var f = new FormCalcMoistureAsh(CalcMode.Ash))
            {
                if (f.ShowDialog(this) == DialogResult.OK &&
                    f.AverageResult.HasValue)
                {
                    CylinderRawAshTB.Text =
                        f.AverageResult.Value.ToString("F2");
                }
            }
        }
        private void AddMaterialToDgv(MaterialInfo info)
        {
            string materialDisplay =
                info.MaterialNo + Environment.NewLine +
                info.MaterialName;
            RawMaterialdgv.Columns[1].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            RawMaterialdgv.Columns[7].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            RawMaterialdgv.Columns[8].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            RawMaterialdgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            string inDate = MaterialTestDateBox.Value.ToString("yyyy.MM.dd");
            if (RawMaterialdgv.Rows.Count >= MAX_ROWS)
            {
                MessageBox.Show($"最多只能新增 {MAX_ROWS} 筆資料");
                return;
            }
            RawMaterialdgv.Rows.Add(
                inDate,
                materialDisplay,
                "",
                info.InUnit,
                info.SampleQty,
                info.InspectUnit,
                "V",
                info.Spec,
                info.Spec,
                "合格",
                "N/A"
            );
        }
        private void RawMaterialNOtb_keyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            e.SuppressKeyPress = true; // 防止嗶聲

            string materialNo = RawMaterialNOtb.Text.Trim();

            var info = MaterialMasterHelper.Get(materialNo);

            if (info == null)
            {
                var r = MessageBox.Show(
                    "查無此物料，是否需要新增？",
                    "查無資料",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (r == DialogResult.Yes)
                {
                    AddMaterial(materialNo);
                }

                return;
            }
            // ===== 找到了 =====
            AddMaterialToDgv(info);
        }
        private void AddMaterial(string materialNo)
        {
            using (var f = new FormAddMaterial(materialNo))
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    MaterialMasterHelper.Add(f.Result);
                    AddMaterialToDgv(f.Result);
                }
            }
        }
        private void CylinderNOBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            string cylinderNo = CylinderNoBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(cylinderNo))
                return;

            // ★ 查詢總表（可能回多筆）
            var result = Page5LookupHelper.SearchByCylinderNo(cylinderNo);

            // ---- 查無資料 ----
            if (!result.Found || result.Rows.Count == 0)
            {
                MessageBox.Show(
                    "總表中查無此單號資料。",
                    "查詢結果",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                ClearPage5();
                return;
            }
            FillHeaderFromLookup(result);
            FillDgvFromLookup(result);

            MessageBox.Show(
                "已帶入該單號最新的檢驗結果。",
                "查詢完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        private void FillHeaderFromLookup(Page5LookupResult result)
        {
            CylinderTestDateBox.Text = result.HeaderValues.GetValueOrDefault("U");
            CylinderReportNOBox.Text = result.HeaderValues.GetValueOrDefault("V");
            CylinderCustmorBox.Text = result.HeaderValues.GetValueOrDefault("Y");
            CYLTypeBox.Text = result.HeaderValues.GetValueOrDefault("AA");
            ReCylinderBox.Text =
                result.HeaderValues.ContainsKey("AB")
                    ? result.HeaderValues["AB"]
                    : "";
            CYLRawMaterialBox.Text = result.HeaderValues.GetValueOrDefault("X");
            string ai = result.HeaderValues.GetValueOrDefault("AI");
            string aj = result.HeaderValues.GetValueOrDefault("AJ");
            string ak = result.HeaderValues.GetValueOrDefault("AK");

            string materialLot = new[] { ai, aj, ak }
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v) && v != "N/A");

            if (!string.IsNullOrWhiteSpace(materialLot))
            {
                // 有原料批號 → 直接帶入
                CYLRawEffTB.Text = materialLot;
            }
            else
            {
                // 沒有原料批號 → 清空並提醒
                CYLRawEffTB.Text = "";

                MessageBox.Show(
                    "此單號尚未填入原料批號。",
                    "原料批號提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }

        }
        private void CYLRawMaterialBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;
            CYLRawEffTB.Clear();
            string materialLot = CYLRawMaterialBox.Text?.Trim();

            // ---- 無效輸入：直接不做事 ----
            if (string.IsNullOrWhiteSpace(materialLot) ||
                materialLot == "-" ||
                materialLot.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            LookupRawMaterialEfficiency(materialLot);
        }
        private void LookupRawMaterialEfficiency(string materialLot)
        {
            string type = CYLTypeBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(type))
            {
                MessageBox.Show(
                    "請先選擇原料種類。",
                    "必要欄位未填",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                CYLTypeBox.Focus(); // 導引使用者
                return;
            }

            var targetTypes = new List<string> { type };

            string totalPath = @"C:\Users\User\Desktop\總表.xlsx";

            try
            {
                using (var wb = new ClosedXML.Excel.XLWorkbook(totalPath))
                {
                    var ws = wb.Worksheet("濾筒");

                    var eff = EfficiencyFinder.FindMinEfficiencyByCarbonLot(
                        ws,
                        materialLot,
                        targetTypes
                    );

                    // ---------- 狀況 1：查無此批號 ----------
                    if (eff == null || eff.Count == 0)
                    {
                        MessageBox.Show(
                            "查無此原料批號的歷史效率資料，請確認批號是否正確。",
                            "查詢結果",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    // ---------- 取出第一個有效效率 ----------
                    string effValueRaw = eff.Values.FirstOrDefault(v =>
                !string.IsNullOrWhiteSpace(v) &&
                v != "-" &&
                !v.Equals("N/A", StringComparison.OrdinalIgnoreCase)
            );

                    // ---------- 狀況 2：有批號，但效率無效 ----------
                    if (string.IsNullOrWhiteSpace(effValueRaw))
                    {
                        MessageBox.Show(
                            "查無此批號",
                            "效率資料提醒",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                        );
                        return;
                    }

                    // ---------- 狀況 3：成功 ----------
                    if (double.TryParse(effValueRaw, out double effValue))
                    {
                        CYLRawEffTB.Text = effValue.ToString("0.0");
                    }
                    else
                    {
                        // 理論上不該發生，保險處理
                        CYLRawEffTB.Text = effValue.ToString("0.0"); ;
                    }
                }
            }
            catch
            {
                // 依你目前的設計哲學：錯誤時不干擾使用者
                return;
            }
        }
    }
}



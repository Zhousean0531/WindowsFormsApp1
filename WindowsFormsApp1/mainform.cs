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

        }
        private void TxtAsh_Click(object sender, EventArgs e)
        {

        }

        private void AddMaterialToDgv(MaterialInfo info)
        {
            string materialDisplay =
                info.MaterialNo + Environment.NewLine +
                info.MaterialName;

            string inDate = MaterialTestDateBox.Value.ToString("yyyy.MM.dd");

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
                MessageBox.Show("查無此料號");
                return;
            }

            AddMaterialToDgv(info);
        }
    }
}



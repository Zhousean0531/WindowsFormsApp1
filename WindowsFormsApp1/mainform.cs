using System;
using System.Collections.Generic;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;


namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        private const int MAX_ROWS = 13;
        private ComboBox _cylinderCustomerComboBox;
        private Label _filterInProcessOrderTypeLabel;
        private ComboBox _filterInProcessOrderTypeBox;

        public Form1()
        {
            InitializeComponent();
            InitializeCylinderCustomerComboBox();
            InitializeFilterInProcessOrderTypeBox();
            this.FilterRawTypeBox.SelectedIndexChanged += new System.EventHandler(this.FilterRawTypeBox_SelectedIndexChanged);
            string today = DateTime.Now.ToString("yyyy.MM.dd");
            CylinderRawMoistureTB.Click += TxtMoisture_Click;
            CylinderRawAshTB.Click += TxtAsh_Click;
            chkMoisture.CheckedChanged += (s, e) => ToggleMoistureUI();
            chkAsh.CheckedChanged += (s, e) => ToggleAshUI();
            ToggleMoistureUI();
            ToggleAshUI();
            InitializeQueryPage();
            InitializeQualityAnalysisLauncher();
            InitializeClearCurrentPageButton();
            InitializeResponsiveLayout();
        }

        private void InitializeCylinderCustomerComboBox()
        {
            if (CylinderCustmorBox == null || CylinderCustmorBox.Parent == null)
                return;

            var template = CylinderCustmorBox;
            var parent = template.Parent;
            int childIndex = parent.Controls.GetChildIndex(template);
            string initialText = string.IsNullOrWhiteSpace(template.Text)
                ? "General"
                : template.Text.Trim();

            _cylinderCustomerComboBox = new ComboBox
            {
                Name = template.Name,
                Location = template.Location,
                Size = template.Size,
                Anchor = template.Anchor,
                Dock = template.Dock,
                Font = template.Font,
                Cursor = template.Cursor,
                TabIndex = template.TabIndex,
                DropDownStyle = ComboBoxStyle.DropDown,
                FormattingEnabled = true,
                Text = initialText
            };
            _cylinderCustomerComboBox.Items.Add("General");

            parent.Controls.Remove(template);
            template.Visible = false;
            template.Name = "CylinderCustmorTextBoxTemplate";

            parent.Controls.Add(_cylinderCustomerComboBox);
            parent.Controls.SetChildIndex(_cylinderCustomerComboBox, childIndex);
        }

        private void InitializeFilterInProcessOrderTypeBox()
        {
            if (FilterInProcessPage == null)
                return;

            if (ControlHelper.Find<ComboBox>(FilterInProcessPage, "FilterInProcessOrderTypeBox") != null)
                return;

            _filterInProcessOrderTypeLabel = new Label
            {
                AutoSize = true,
                Location = new Point(306, 300),
                Name = "FilterInProcessOrderTypeLabel",
                Text = "訂單類型"
            };

            _filterInProcessOrderTypeBox = new ComboBox
            {
                Cursor = Cursors.Default,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FormattingEnabled = true,
                Location = new Point(386, 295),
                Name = "FilterInProcessOrderTypeBox",
                Size = new Size(120, 28),
                TabIndex = 21
            };

            _filterInProcessOrderTypeBox.Items.AddRange(new object[]
            {
                "一般",
                "台積堆疊式"
            });
            _filterInProcessOrderTypeBox.SelectedIndex = 0;

            FilterInProcessPage.Controls.Add(_filterInProcessOrderTypeLabel);
            FilterInProcessPage.Controls.Add(_filterInProcessOrderTypeBox);
        }

        private void SetCylinderCustomerText(string value)
        {
            string text = (value ?? string.Empty).Trim();
            var combo = _cylinderCustomerComboBox
                ?? ControlHelper.Find<ComboBox>(CylinderPage, "CylinderCustmorBox");

            if (combo != null)
            {
                if (!string.IsNullOrWhiteSpace(text))
                {
                    bool exists = false;
                    foreach (var item in combo.Items)
                    {
                        if (string.Equals(item?.ToString(), text, StringComparison.OrdinalIgnoreCase))
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (!exists)
                        combo.Items.Add(text);
                }

                combo.Text = text;
                return;
            }

            CylinderCustmorBox.Text = text;
        }

        private void ClearPage5()
        {
            CylinderBox.Rows.Clear();
            CylinderReportNOBox.Clear();
            SetCylinderCustomerText("");
            ReCylinderBox.Clear();
            CYLTypeBox.Text = "";
            CYLFilterNameBox.Clear();
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
            FilterRawBatchNOBox.Text = $"B-{dt}-";
        }
        private void CylinderRawArriveDateBox_ValueChanged(object sender, EventArgs e)
        {
            var dt = CylinderRawArriveDateBox.Value.ToString("yyyyMMdd");
            CylinderRawLotFullTB.Text = $"B-{dt}-";
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
            if (!DbBootstrap.Reconnect())
                return;

            var currentTab = tabControl1.SelectedTab;
            
            var tab = tabControl1.SelectedTab;

            // 只有 Page1 / Page4 才需要 mesh
            if (tab.Name == "FilterRawPage" || tab.Name == "CylinderRawPage")
            {
                var dgv =
                    ControlHelper.Find<DataGridView>(tab, "FilterRawParticleSizeBox")
                    ?? ControlHelper.Find<DataGridView>(tab, "CylinderRawMeshBox");

                if (!MeshGridHelper.RecalculateAndFill(dgv))
                {
                    MessageBox.Show("粒徑資料不完整，請確認");
                    return;
                }
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
                case "QueryPage":
                    RunQcQuery();
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
                "Y",
                info.Spec,
                info.Spec,
                "合格",
                "N/A",
                ""
            );
        }
        private void RawMaterialNOtb_keyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
                return;

            e.SuppressKeyPress = true; // 防止嗶聲

            // 正規化使用者輸入為大寫（與 _map 儲存一致）
            string materialNo = RawMaterialNOtb.Text.Trim().ToUpper();

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
                    "資料庫中查無此單號資料。",
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
            // ===== 基本欄位 =====
            CylinderTestDateBox.Text = result.HeaderValues.GetValueOrDefault("TestDate");
            CylinderReportNOBox.Text = result.HeaderValues.GetValueOrDefault("ReportNo");
            CylinderNoBox.Text = result.HeaderValues.GetValueOrDefault("CylinderNo");
            SetCylinderCustomerText(result.HeaderValues.GetValueOrDefault("Customer"));

            // 原料種類 (MA / MB / MC)
            CYLTypeBox.Text = result.HeaderValues.GetValueOrDefault("RawMaterialType");

            // 原料名稱
            CYLFilterNameBox.Text = result.HeaderValues.GetValueOrDefault("FilterType");

            // 再生筒號
            ReCylinderBox.Text = result.HeaderValues.GetValueOrDefault("ReCylinderNo");

            // ===== 原料批號 =====
            string carbonLot = result.HeaderValues.GetValueOrDefault("CarbonLot");

            if (!string.IsNullOrWhiteSpace(carbonLot))
            {
                CYLRawMaterialBox.Text = carbonLot;
            }
            else
            {
                CYLRawMaterialBox.Text = "";

                MessageBox.Show(
                    "此單號尚未填入原料批號。",
                    "原料批號提醒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }

            // ===== 原料效率 =====
            string efficiency = result.HeaderValues.GetValueOrDefault("Efficiency");

            if (!string.IsNullOrWhiteSpace(efficiency))
            {
                CYLRawEffTB.Text = efficiency;
            }
            else
            {
                CYLRawEffTB.Text = "";

                MessageBox.Show(
                    "此單號尚未填入原料效率。",
                    "效率提醒",
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

                CYLTypeBox.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(materialLot))
            {
                MessageBox.Show(
                    "請先輸入原料批號。",
                    "必要欄位未填",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );

                CYLRawMaterialBox.Focus();
                return;
            }
            
            try
            {

                string effValueRaw = EfficiencyFinder.FindMinEfficiencyByCarbonLot(
                     materialLot,
                     type
                 );

                if (string.IsNullOrWhiteSpace(effValueRaw))
                {
                    MessageBox.Show(
                        "查無此原料批號的歷史效率資料，請確認批號或原料種類是否正確。",
                        "查詢結果",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                if (effValueRaw == "-" ||
                    effValueRaw.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(
                        "此原料批號有歷史資料，但效率資料無效。",
                        "效率資料提醒",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    return;
                }

                if (double.TryParse(effValueRaw, out double effValue))
                {
                    CYLRawEffTB.Text = effValue.ToString("0.0");
                }
                else
                {
                    MessageBox.Show(
                        "效率資料格式異常，無法轉換為數字。",
                        "效率資料錯誤",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "查詢原料效率時發生錯誤：" + ex.Message,
                    "查詢錯誤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            string importKind = AskExcelImportKind();
            if (string.IsNullOrWhiteSpace(importKind))
                return;

            using (var ofd = new OpenFileDialog())
            {
                ofd.Title = $"選擇 {importKind} 匯入 Excel";
                ofd.Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";

                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                try
                {
                    int count = ImportExcelByKind(ofd.FileName, importKind);

                    MessageBox.Show($"匯入完成，共 {count} 批");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        "匯入失敗：" + ex.Message,
                        "匯入錯誤",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
        }

        private string AskExcelImportKind()
        {
            using (var dialog = new Form())
            using (var label = new Label())
            using (var combo = new ComboBox())
            using (var okButton = new Button())
            using (var cancelButton = new Button())
            {
                dialog.Text = "選擇匯入資料";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ClientSize = new Size(360, 145);
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.ShowInTaskbar = false;

                label.AutoSize = true;
                label.Location = new Point(24, 24);
                label.Text = "今天想匯入哪一種資料？";

                combo.DropDownStyle = ComboBoxStyle.DropDownList;
                combo.Location = new Point(24, 55);
                combo.Size = new Size(310, 28);
                combo.Items.AddRange(new object[]
                {
                    "濾網成品",
                    "濾網半成品",
                    "濾網原料",
                    "濾筒原料",
                    "濾筒成品",
                    "物料"
                });
                combo.SelectedIndex = 0;

                okButton.Text = "確定";
                okButton.DialogResult = DialogResult.OK;
                okButton.Location = new Point(168, 101);
                okButton.Size = new Size(80, 30);

                cancelButton.Text = "取消";
                cancelButton.DialogResult = DialogResult.Cancel;
                cancelButton.Location = new Point(254, 101);
                cancelButton.Size = new Size(80, 30);

                dialog.Controls.Add(label);
                dialog.Controls.Add(combo);
                dialog.Controls.Add(okButton);
                dialog.Controls.Add(cancelButton);
                dialog.AcceptButton = okButton;
                dialog.CancelButton = cancelButton;

                return dialog.ShowDialog(this) == DialogResult.OK
                    ? combo.Text
                    : null;
            }
        }

        private int ImportExcelByKind(string fileName, string importKind)
        {
            switch (importKind)
            {
                case "濾網原料":
                    return Page1ExcelImporter.ImportFromExcel(fileName, importKind);
                case "濾網半成品":
                    return Page2ExcelImporter.ImportFromExcel(fileName, importKind);
                case "濾網成品":
                    return Page3ExcelImporter.ImportFromExcel(fileName, importKind);
                case "濾筒原料":
                    return Page4ExcelImporter.ImportFromExcel(fileName, importKind);
                case "濾筒成品":
                    return Page5ExcelImporter.ImportFromExcel(fileName, importKind);
                case "物料":
                    return Page6ExcelImporter.ImportFromExcel(fileName, importKind);
                default:
                    throw new InvalidOperationException("不支援的匯入資料種類：" + importKind);
            }
        }
    }
}




using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Query;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private const int QueryLeft = 32;
        private const int QueryGridTop = 306;
        private const int QueryRightMargin = 36;
        private const int QueryBottomMargin = 40;

        private TabPage QueryPage;
        private DateTimePicker QueryStartDateBox;
        private DateTimePicker QueryEndDateBox;
        private ComboBox QueryKindBox;
        private TextBox QueryRawMaterialBox;
        private TextBox QuerySemiProductBox;
        private TextBox QuerySemiGsmBox;
        private TextBox QuerySemiMaterialNoBox;
        private TextBox QueryProductNoBox;
        private TextBox QueryMaterialBox;
        private Button QuerySearchButton;
        private Button QueryExportButton;
        private DataGridView QueryResultGrid;
        private DataTable _queryResultTable;

        private void InitializeQueryPage()
        {
            QueryPage = new TabPage
            {
                Name = "QueryPage",
                Text = "查詢",
                BackColor = Color.White,
                UseVisualStyleBackColor = false
            };
            QueryPage.Resize += QueryPage_Resize;

            var title = new Label
            {
                Text = "QC Report 查詢",
                Font = QueryFont(16F, FontStyle.Bold),
                Location = new Point(32, 22),
                AutoSize = true
            };

            QueryStartDateBox = BuildDateBox(132, 78);
            QueryEndDateBox = BuildDateBox(330, 78);

            QueryKindBox = new ComboBox
            {
                Name = "QueryKindBox",
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = QueryFont(11F),
                Location = new Point(132, 124),
                Size = new Size(170, 27)
            };
            QueryKindBox.Items.AddRange(new object[] { "原料", "半成品", "成品", "物料" });
            QueryKindBox.SelectedIndexChanged += QueryKindBox_SelectedIndexChanged;

            QueryRawMaterialBox = BuildConditionTextBox("QueryRawMaterialBox", 132, 172);
            QuerySemiProductBox = BuildConditionTextBox("QuerySemiProductBox", 132, 218);
            QuerySemiGsmBox = BuildConditionTextBox("QuerySemiGsmBox", 430, 218);
            QuerySemiMaterialNoBox = BuildConditionTextBox("QuerySemiMaterialNoBox", 728, 218);
            QueryProductNoBox = BuildConditionTextBox("QueryProductNoBox", 132, 264);
            QueryMaterialBox = BuildConditionTextBox("QueryMaterialBox", 430, 264);

            QuerySearchButton = new Button
            {
                Text = "查詢",
                Font = QueryFont(12F),
                Location = new Point(706, 76),
                Size = new Size(100, 42),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            QuerySearchButton.Click += QuerySearchButton_Click;

            QueryExportButton = new Button
            {
                Text = "匯出",
                Font = QueryFont(12F),
                Location = new Point(826, 76),
                Size = new Size(100, 42),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            QueryExportButton.Click += QueryExportButton_Click;

            QueryResultGrid = new DataGridView
            {
                Name = "QueryResultGrid",
                Location = new Point(QueryLeft, QueryGridTop),
                Size = new Size(900, 240),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                ScrollBars = ScrollBars.Both,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            QueryPage.Controls.Add(title);
            QueryPage.Controls.Add(BuildLabel("查詢期間：", 48, 82));
            QueryPage.Controls.Add(QueryStartDateBox);
            QueryPage.Controls.Add(BuildLabel("~", 300, 83));
            QueryPage.Controls.Add(QueryEndDateBox);
            QueryPage.Controls.Add(BuildLabel("查詢種類：", 48, 128));
            QueryPage.Controls.Add(QueryKindBox);
            QueryPage.Controls.Add(BuildLabel("原料種類：", 48, 176));
            QueryPage.Controls.Add(QueryRawMaterialBox);
            QueryPage.Controls.Add(BuildLabel("半成品種類：", 32, 222));
            QueryPage.Controls.Add(QuerySemiProductBox);
            QueryPage.Controls.Add(BuildLabel("半成品克重：", 330, 222));
            QueryPage.Controls.Add(QuerySemiGsmBox);
            QueryPage.Controls.Add(BuildLabel("半成品料號：", 628, 222));
            QueryPage.Controls.Add(QuerySemiMaterialNoBox);
            QueryPage.Controls.Add(BuildLabel("成品料號：", 48, 268));
            QueryPage.Controls.Add(QueryProductNoBox);
            QueryPage.Controls.Add(BuildLabel("物料料號/名稱：", 314, 268));
            QueryPage.Controls.Add(QueryMaterialBox);
            QueryPage.Controls.Add(QuerySearchButton);
            QueryPage.Controls.Add(QueryExportButton);
            QueryPage.Controls.Add(QueryResultGrid);

            tabControl1.Controls.Add(QueryPage);
            SetQueryConditionState();
            LayoutQueryPage();
        }

        private Font QueryFont(float size, FontStyle style = FontStyle.Regular)
        {
            return new Font("Microsoft JhengHei UI", size, style);
        }

        private void QueryPage_Resize(object sender, EventArgs e)
        {
            LayoutQueryPage();
        }

        private void LayoutQueryPage()
        {
            if (QueryPage == null || QueryResultGrid == null)
                return;

            int gridWidth = Math.Max(520, QueryPage.ClientSize.Width - QueryLeft - QueryRightMargin);
            int gridHeight = Math.Max(150, QueryPage.ClientSize.Height - QueryGridTop - QueryBottomMargin);
            QueryResultGrid.Bounds = new Rectangle(QueryLeft, QueryGridTop, gridWidth, gridHeight);

            if (QuerySearchButton != null)
                QuerySearchButton.Left = Math.Max(560, QueryPage.ClientSize.Width - 284);

            if (QueryExportButton != null)
                QueryExportButton.Left = Math.Max(680, QueryPage.ClientSize.Width - 164);
        }

        private DateTimePicker BuildDateBox(int x, int y)
        {
            var picker = new DateTimePicker
            {
                Font = QueryFont(11F),
                Format = DateTimePickerFormat.Custom,
                CustomFormat = " ",
                Location = new Point(x, y),
                Size = new Size(150, 27),
                Tag = false
            };

            picker.ValueChanged += QueryDateBox_ValueChanged;
            picker.KeyDown += QueryDateBox_KeyDown;

            return picker;
        }

        private TextBox BuildConditionTextBox(string name, int x, int y)
        {
            return new TextBox
            {
                Name = name,
                Font = QueryFont(11F),
                Location = new Point(x, y),
                Size = new Size(170, 27),
                Enabled = false
            };
        }

        private Label BuildLabel(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Font = QueryFont(11F),
                Location = new Point(x, y),
                AutoSize = true
            };
        }

        private void QueryKindBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetQueryConditionState();
        }

        private void SetQueryConditionState()
        {
            string kind = QueryKindBox?.Text ?? "";

            bool isRaw = kind == "原料";
            bool isSemi = kind == "半成品";
            bool isProduct = kind == "成品";
            bool isMaterial = kind == "物料";

            SetBoxEnabled(QueryRawMaterialBox, isRaw);
            SetBoxEnabled(QuerySemiProductBox, isSemi);
            SetBoxEnabled(QuerySemiGsmBox, isSemi);
            SetBoxEnabled(QuerySemiMaterialNoBox, isSemi);
            SetBoxEnabled(QueryProductNoBox, isProduct);
            SetBoxEnabled(QueryMaterialBox, isMaterial);

            QuerySearchButton.Enabled = !string.IsNullOrWhiteSpace(kind);
        }

        private void SetBoxEnabled(TextBox box, bool enabled)
        {
            if (box == null)
                return;

            box.Enabled = enabled;
            box.BackColor = enabled ? Color.White : SystemColors.Control;

            if (!enabled)
                box.Clear();
        }

        private void QuerySearchButton_Click(object sender, EventArgs e)
        {
            RunQcQuery();
        }

        private void RunQcQuery()
        {
            if (QueryKindBox.SelectedIndex < 0)
            {
                MessageBox.Show("請先選擇查詢種類");
                return;
            }

            try
            {
                var criteria = new QcQueryCriteria
                {
                    QueryKind = QueryKindBox.Text,
                    StartDate = GetQueryDate(QueryStartDateBox),
                    EndDate = GetQueryDate(QueryEndDateBox),
                    RawMaterialType = QueryRawMaterialBox.Text.Trim(),
                    SemiProductType = QuerySemiProductBox.Text.Trim(),
                    SemiProductGsm = QuerySemiGsmBox.Text.Trim(),
                    SemiProductMaterialNo = QuerySemiMaterialNoBox.Text.Trim(),
                    ProductNo = QueryProductNoBox.Text.Trim(),
                    MaterialKeyword = QueryMaterialBox.Text.Trim()
                };

                _queryResultTable = QcQueryRepository.Query(criteria);
                QueryResultGrid.DataSource = _queryResultTable;
                ConfigureQueryResultGrid();
                QueryExportButton.Enabled = _queryResultTable.Rows.Count > 0;

                MessageBox.Show($"查詢完成，共 {_queryResultTable.Rows.Count} 筆");
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);
            }
        }

        private void QueryDateBox_ValueChanged(object sender, EventArgs e)
        {
            var picker = sender as DateTimePicker;
            if (picker == null)
                return;

            picker.CustomFormat = "yyyy-MM-dd";
            picker.Tag = true;
        }

        private void QueryDateBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Delete && e.KeyCode != Keys.Back)
                return;

            var picker = sender as DateTimePicker;
            if (picker == null)
                return;

            picker.CustomFormat = " ";
            picker.Tag = false;
            e.SuppressKeyPress = true;
        }

        private DateTime? GetQueryDate(DateTimePicker picker)
        {
            if (picker == null)
                return null;

            return picker.Tag is bool hasValue && hasValue
                ? picker.Value.Date
                : (DateTime?)null;
        }

        private void QueryExportButton_Click(object sender, EventArgs e)
        {
            QueryExcelExporter.Export(_queryResultTable);
        }

        private void ConfigureQueryResultGrid()
        {
            if (QueryResultGrid == null)
                return;

            QueryResultGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            QueryResultGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            QueryResultGrid.ScrollBars = ScrollBars.Both;
            QueryResultGrid.RowHeadersWidth = 28;

            foreach (DataGridViewColumn col in QueryResultGrid.Columns)
            {
                col.Width = 140;
                col.MinimumWidth = 80;
                col.Resizable = DataGridViewTriState.True;
            }
        }
    }
}

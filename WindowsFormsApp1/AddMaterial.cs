
using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class FormAddMaterial : Form
    {
        // ===== 對外結果 =====
        public MaterialInfo Result { get; private set; }

        // ===== TextBox =====
        TextBox tbMaterialNo;
        TextBox tbMaterialName;
        TextBox tbInUnit;
        TextBox tbSampleQty;
        TextBox tbInspectUnit;
        TextBox tbSpec;

        Button btnOk;
        Button btnCancel;

        public FormAddMaterial(string materialNo)
        {
            Text = "新增料號";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Width = 360;
            Height = 360;

            InitControls(materialNo);
        }

        // ================= UI 建立 =================
        private void InitControls(string materialNo)
        {
            int labelX = 20;
            int tbX = 120;
            int y = 20;
            int gap = 35;

            // 料號
            CreateLabel("料號", labelX, y);
            tbMaterialNo = CreateTextBox(tbX, y);
            tbMaterialNo.Text = materialNo;
            tbMaterialNo.ReadOnly = true;
            y += gap;

            // 物料名稱
            CreateLabel("物料名稱", labelX, y);
            tbMaterialName = CreateTextBox(tbX, y);
            y += gap;

            // 進貨單位
            CreateLabel("進貨單位", labelX, y);
            tbInUnit = CreateTextBox(tbX, y);
            y += gap;

            // 抽檢數量
            CreateLabel("抽檢數量", labelX, y);
            tbSampleQty = CreateTextBox(tbX, y);
            y += gap;

            // 檢驗單位
            CreateLabel("檢驗單位", labelX, y);
            tbInspectUnit = CreateTextBox(tbX, y);
            y += gap;

            // 規格值
            CreateLabel("規格值", labelX, y);
            tbSpec = CreateTextBox(tbX, y);
            y += gap + 10;

            // Buttons
            btnOk = new Button
            {
                Text = "確認",
                Left = 80,
                Top = y,
                Width = 80
            };
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Left = 180,
                Top = y,
                Width = 80
            };
            btnCancel.Click += (s, e) => Close();

            Controls.Add(btnOk);
            Controls.Add(btnCancel);
        }

        private Label CreateLabel(string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Left = x,
                Top = y + 4,
                Width = 90
            };
            Controls.Add(lbl);
            return lbl;
        }

        private TextBox CreateTextBox(int x, int y)
        {
            var tb = new TextBox
            {
                Left = x,
                Top = y,
                Width = 180
            };
            Controls.Add(tb);
            return tb;
        }

        // ================= 確認按鈕 =================
        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(tbMaterialName.Text))
            {
                MessageBox.Show("請填寫物料名稱");
                return;
            }

            Result = new MaterialInfo
            {
                MaterialNo = tbMaterialNo.Text.Trim().ToUpper(),
                MaterialName = tbMaterialName.Text.Trim(),
                InUnit = tbInUnit.Text.Trim(),
                SampleQty = tbSampleQty.Text.Trim(),
                InspectUnit = tbInspectUnit.Text.Trim(),
                Spec = tbSpec.Text.Trim()
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(tbMaterialName.Text) ||
                string.IsNullOrWhiteSpace(tbInUnit.Text) ||
                string.IsNullOrWhiteSpace(tbSampleQty.Text) ||
                string.IsNullOrWhiteSpace(tbInspectUnit.Text) ||
                string.IsNullOrWhiteSpace(tbSpec.Text))
            {
                MessageBox.Show("請填寫所有欄位", "資料不完整",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
    }
}

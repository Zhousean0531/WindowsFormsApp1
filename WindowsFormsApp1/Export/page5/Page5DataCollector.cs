using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

public static class Page5DataCollector
{
    public static Page5ExportData Collect(TabPage tab)
    {
        // ===== 建立基本資料 =====
        string userName = Environment.UserName;
        var data = new Page5ExportData
        {
            TestDate = ControlHelper.GetText(tab, "CylinderTestDateBox"),
            ReportNo = ControlHelper.GetText(tab, "CylinderReportNOBox"),
            CylinderNo = ControlHelper.GetText(tab, "CylinderNoBox"),
            Customer = ControlHelper.GetText(tab, "CylinderCustmorBox"),
            FilterType = ControlHelper.GetText(tab, "CYLTypeBox"),
            ReCylinderNo = ControlHelper.GetText(tab, "ReCylinderBox"),
            CarbonLot = ControlHelper.GetText(tab, "CYLRawMaterialBox"),
            Rows = new List<Page5RowData>(),
            UserName = userName
        };

        // ===== 基本欄位防呆 =====
        var missingFields = new List<string>();

        if (string.IsNullOrWhiteSpace(data.TestDate))
            missingFields.Add("測試日期");
        if (string.IsNullOrWhiteSpace(data.ReportNo))
            missingFields.Add("報告編號");
        if (string.IsNullOrWhiteSpace(data.CylinderNo))
            missingFields.Add("生產單號");
        if (string.IsNullOrWhiteSpace(data.Customer))
            missingFields.Add("客戶");
        if (string.IsNullOrWhiteSpace(data.FilterType))
            missingFields.Add("原料種類");

        if (missingFields.Any())
        {
            MessageBox.Show(
                "以下基本欄位尚未填寫：\n" + string.Join("、", missingFields),
                "資料未完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return null;
        }

        data.MaterialNo = ResolveMaterialNo(data.FilterType);
        if (data.MaterialNo == null)
            return null;

        // ===== 取得 DataGridView =====
        var dgv = tab.Controls.Find("CylinderBox", true)
                              .FirstOrDefault() as DataGridView;

        if (dgv == null)
        {
            MessageBox.Show(
                "找不到量測資料表，請確認表單是否正常載入。",
                "系統錯誤",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return null;
        }

        var rows = dgv.Rows.Cast<DataGridViewRow>()
                           .Where(r => !r.IsNewRow)
                           .ToList();

        // ===== DGV 為空防呆 =====
        if (rows.Count == 0)
        {
            MessageBox.Show(
                "尚未輸入任何量測資料，請先填寫表格內容。",
                "資料未完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return null;
        }

        // ===== 16 筆一組檢查 =====
        if (rows.Count % 16 != 0)
        {
            MessageBox.Show(
                "資料筆數異常（必須為 16 的倍數）",
                "資料結構錯誤",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
            return null;
        }

        // ===== 欄位定義 =====
        string[] requiredControlColumns =
        {
            "CYL_Particle_In",
            "CYL_Particle_out",
            "CYL_IPA_in",
            "CYL_IPA_out",
            "CYL_Acetone_In",
            "CYL_Acetone_out",
            "CYL_Nontarget_in",
            "CYL_Nontarget_out",
            "CYL_Pressure_Drop"
        };

        string sampleSnCol = "CYLSN";
        string sampleWeightCol = "CYLWeight";

        Dictionary<int, string> controlValues = null;

        // ===== 主迴圈 =====
        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            bool isControlRow = i % 16 == 0;

            // ---- 控制列檢查 ----
            if (isControlRow)
            {
                int groupIndex = i / 16 + 1;
                var missingCols = new List<string>();

                foreach (var col in requiredControlColumns)
                {
                    var v = r.Cells[col]?.Value?.ToString();
                    if (string.IsNullOrWhiteSpace(v))
                        missingCols.Add(col);
                }

                if (missingCols.Any())
                {
                    MessageBox.Show(
                        $"第 {groupIndex} 組控制列資料未填完整：\n" +
                        string.Join("、", missingCols),
                        "控制列資料錯誤",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return null;
                }

                controlValues = BuildControlValues(r);
            }
            else
            {
                // ---- 樣本列檢查（SN / Weight 成對） ----
                string sn = r.Cells[sampleSnCol]?.Value?.ToString();
                string weight = r.Cells[sampleWeightCol]?.Value?.ToString();

                bool snHas = !string.IsNullOrWhiteSpace(sn);
                bool weightHas = !string.IsNullOrWhiteSpace(weight);

                if (snHas ^ weightHas)
                {
                    int displayRow = i + 1;

                    MessageBox.Show(
                        $"第 {displayRow} 列資料不完整：\n" +
                        "樣本編號與重量必須同時填寫。",
                        "樣本資料錯誤",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return null;
                }
            }

            // ---- 建立 RowData ----
            var rowData = new Page5RowData
            {
                SN = r.Cells["CYLSN"]?.Value?.ToString(),
                Weight = r.Cells["CYLWeight"]?.Value?.ToString(),
                ControlValues = new Dictionary<int, string>(controlValues)
            };

            data.Rows.Add(rowData);
        }

        return data;
    }

    // ===== 建立控制列資料 =====
    private static Dictionary<int, string> BuildControlValues(DataGridViewRow r)
    {
        string Get(string name) => r.Cells[name]?.Value?.ToString() ?? "";

        return new Dictionary<int, string>
        {
            [38] = Get("CYL_Particle_In"),
            [39] = Get("CYL_Particle_out"),
            [40] = DiffUtil.GetDiff(Get("CYL_Particle_out"), Get("CYL_Particle_In")),

            [41] = Get("CYL_IPA_in"),
            [42] = Get("CYL_IPA_out"),
            [43] = DiffUtil.GetDiff(Get("CYL_IPA_out"), Get("CYL_IPA_in")),

            [44] = Get("CYL_Acetone_In"),
            [45] = Get("CYL_Acetone_out"),
            [46] = DiffUtil.GetDiff(Get("CYL_Acetone_out"), Get("CYL_Acetone_In")),

            [47] = Get("CYL_Nontarget_in"),
            [48] = Get("CYL_Nontarget_out"),
            [49] = DiffUtil.GetDiff(Get("CYL_Nontarget_out"), Get("CYL_Nontarget_in")),

            [50] = DiffUtil.GetSumDiff(
                Get("CYL_IPA_out"), Get("CYL_Acetone_out"), Get("CYL_Nontarget_out"),
                Get("CYL_IPA_in"), Get("CYL_Acetone_In"), Get("CYL_Nontarget_in")
            ),

            [54] = Get("CYL_Pressure_Drop")
        };
    }

    private static string ResolveMaterialNo(string filterType)
    {
        string text = (filterType ?? "").Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(text))
            return null;

        var materialNos = new List<string>();

        foreach (string token in text.Split(new[] { '+', '/', ',', '，', '、' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string type = token.Trim().ToUpperInvariant();
            string materialNo = ResolveSingleMaterialNo(type);

            if (materialNo == null)
                return null;

            if (!materialNos.Contains(materialNo))
                materialNos.Add(materialNo);
        }

        if (materialNos.Count == 0)
        {
            string materialNo = ResolveSingleMaterialNo(text);
            if (materialNo == null)
                return null;

            materialNos.Add(materialNo);
        }

        return string.Join("+", materialNos);
    }

    private static string ResolveSingleMaterialNo(string type)
    {
        if (type == "ACID")
            return "11A0C00Y000002";

        if (type == "DMS")
            return "11D0S00Y000002";

        if (type == "MA")
            return AskMaMaterialNo();

        if (type == "MB" || type == "BASE")
            return "11B0B00Y000002";

        if (type == "MC" || type == "TOC")
            return "11T0C00Y000002";

        MessageBox.Show(
            $"無法判斷原料種類「{type}」對應的料號。",
            "原料種類錯誤",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        );

        return null;
    }

    private static string AskMaMaterialNo()
    {
        string selected = null;

        using (var form = new Form())
        using (var label = new Label())
        using (var acidButton = new Button())
        using (var dmsButton = new Button())
        using (var cancelButton = new Button())
        {
            form.Text = "選擇 MA 料號";
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
            form.ShowIcon = false;
            form.ClientSize = new Size(360, 145);

            label.AutoSize = false;
            label.Text = "原料種類為 MA，請選擇要寫入 P5_Batch 的料號：";
            label.Location = new Point(18, 18);
            label.Size = new Size(320, 42);

            acidButton.Text = "ACID";
            acidButton.Location = new Point(35, 82);
            acidButton.Size = new Size(85, 34);
            acidButton.Click += (s, e) =>
            {
                selected = "11A0C00Y000002";
                form.DialogResult = DialogResult.OK;
            };

            dmsButton.Text = "DMS";
            dmsButton.Location = new Point(137, 82);
            dmsButton.Size = new Size(85, 34);
            dmsButton.Click += (s, e) =>
            {
                selected = "11D0S00Y000002";
                form.DialogResult = DialogResult.OK;
            };

            cancelButton.Text = "取消";
            cancelButton.Location = new Point(239, 82);
            cancelButton.Size = new Size(85, 34);
            cancelButton.DialogResult = DialogResult.Cancel;

            form.Controls.Add(label);
            form.Controls.Add(acidButton);
            form.Controls.Add(dmsButton);
            form.Controls.Add(cancelButton);
            form.AcceptButton = acidButton;
            form.CancelButton = cancelButton;

            return form.ShowDialog() == DialogResult.OK ? selected : null;
        }
    }
}

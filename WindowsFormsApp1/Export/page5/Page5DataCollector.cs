using System;
using System.Collections.Generic;
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
}

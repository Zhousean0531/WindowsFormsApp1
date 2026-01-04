using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class Page5DataCollector
{
    public static Page5ExportData Collect(TabPage tab)
    {
        if (!ValidationHelper.CheckRequiredFields(tab))
        {
            MessageBox.Show("請填寫所有欄位再執行");
            return null;
        }

        var data = new Page5ExportData
        {
            TestDate = ControlHelper.GetText(tab, "CylinderTestDateBox"),
            ReportNo = ControlHelper.GetText(tab, "CylinderReportNOBox"),
            CylinderNo = ControlHelper.GetText(tab, "CylinderNOBox"),
            Customer = ControlHelper.GetText(tab, "CylinderCustmorBox"),
            FilterType = ControlHelper.GetText(tab, "CYLTypeBox"),
            ReCylinderNo = ControlHelper.GetText(tab, "ReCylinderNOBox"),
            CarbonLot = ControlHelper.GetText(tab, "CYLMaterialSNBox"),
            Rows = new List<Page5RowData>()
        };

        var dgv = tab.Controls.Find("CylinderBox", true)
                              .FirstOrDefault() as DataGridView;

        if (dgv == null) return null;

        var rows = dgv.Rows.Cast<DataGridViewRow>()
                           .Where(r => !r.IsNewRow)
                           .ToList();

        if (rows.Count % 16 != 0)
        {
            MessageBox.Show("資料筆數異常（必須為 16 的倍數）");
            return null;
        }

        Dictionary<int, string> controlValues = null;

        for (int i = 0; i < rows.Count; i++)
        {
            var r = rows[i];
            bool isControlRow = i % 16 == 0;

            if (isControlRow)
            {
                controlValues = BuildControlValues(r);
            }

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

    private static Dictionary<int, string> BuildControlValues(DataGridViewRow r)
    {
        string Get(string name) => r.Cells[name]?.Value?.ToString() ?? "";

        return new Dictionary<int, string>
        {
            [35] = Get("CYL_Particle_In"),
            [36] = Get("CYL_Particle_out"),
            [37] = DiffUtil.GetDiff(Get("CYL_Particle_out"), Get("CYL_Particle_In")),

            [38] = Get("CYL_IPA_in"),
            [39] = Get("CYL_IPA_out"),
            [40] = DiffUtil.GetDiff(Get("CYL_IPA_out"), Get("CYL_IPA_in")),

            [41] = Get("CYL_Acetone_In"),
            [42] = Get("CYL_Acetone_out"),
            [43] = DiffUtil.GetDiff(Get("CYL_Acetone_out"), Get("CYL_Acetone_In")),

            [44] = Get("CYL_Nontarget_in"),
            [45] = Get("CYL_Nontarget_out"),
            [46] = DiffUtil.GetDiff(Get("CYL_Nontarget_out"), Get("CYL_Nontarget_in")),

            [47] = DiffUtil.GetSumDiff(
                Get("CYL_IPA_out"), Get("CYL_Acetone_out"), Get("CYL_Nontarget_out"),
                Get("CYL_IPA_in"), Get("CYL_Acetone_In"), Get("CYL_Nontarget_in")
            ),

            [51] = Get("CYL_Pressure_Drop")
        };
    }
}

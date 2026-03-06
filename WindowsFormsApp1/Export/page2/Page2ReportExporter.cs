using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Page2;
using static SignatureHelper;
using Excel = Microsoft.Office.Interop.Excel;

public static class Page2ReportExporter
{
    public static void Export(P2Batch batch)
    {
        if (batch == null || batch.GasTests.Count == 0)
            return;

        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel (*.xlsx)|*.xlsx";

            var firstGas = batch.GasTests.First();

            sfd.FileName =
                $"{batch.ReportNo}_{batch.Material}_{batch.TargetGsm}_{batch.WorkOrder}({DateTime.Parse(batch.TestDate):MMdd}生產)_{firstGas.GasName} .xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            string folder = Path.GetDirectoryName(sfd.FileName);

            Excel.Application app = null;

            try
            {
                app = new Excel.Application();
                app.Visible = false;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                foreach (var gas in batch.GasTests)
                {
                    Export_Report(app, folder, batch, gas);
                }

                foreach (var gas in batch.GasTests)
                {
                    Export_Helper(app, folder, batch, gas);
                }
            }
            finally
            {
                app?.Quit();
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }
    }
    private static void Export_Report(Excel.Application app,string folder,P2Batch batch,P2GasTest gas)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_SemiFinished_Template.xlsx");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_SemiFinished_Template.xlsx");
            return;
        }

        string fileName =
            $"{batch.ReportNo}_{batch.Material}_{batch.TargetGsm}_{batch.WorkOrder}_{gas.GasName} ({DateTime.Parse(batch.TestDate):MMdd}生產).xlsx";

        string savePath = Path.Combine(folder, fileName);

        File.Copy(templatePath, savePath, true);

        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            wb = app.Workbooks.Open(savePath);
            ws = wb.Sheets["濾網半成品報告"];

            var selectedSample = gas.Samples.FirstOrDefault(s => s.IsSelected);
            if (selectedSample == null)
                return;

            int idx = gas.Samples.IndexOf(selectedSample);

            var weights = gas.Samples.Select(s => s.Weight ?? 0).ToList();
            var drops = gas.Samples.Select(s => s.PressureDrop ?? 0).ToList();

            DateTime prodDt = DateTime.Parse(batch.ProductionDate);

            string l1Text =
                $"{DateTime.Parse(batch.TestDate):MM.dd} {batch.Material} {weights[idx]}gsm ({drops[idx]}Pa) -{prodDt:MMdd}生產";

            ws.Range["C6"].Value = batch.ReportNo;
            ws.Range["F7"].Value = batch.Material;
            ws.Range["F8"].Value = batch.FilterSize;
            ws.Range["F9"].Value = batch.WorkOrder;
            ws.Range["H6"].Value = batch.TestDate;

            WriteTextWithSubscriptNumbers(ws.Range["F11"], gas.GasName);

            ws.Range["H10"].Value = weights[idx];
            ws.Range["F12"].Value = gas.Concentration;
            ws.Range["F13"].Value = batch.WindSpeed;
            ws.Range["F14"].Value = drops[idx];
            ws.Range["F16"].Value = selectedSample.Efficiencies.FirstOrDefault();

            ws.Range["L1"].Value = l1Text;

            int startRow = 2;

            for (int i = 0; i < selectedSample.Efficiencies.Count; i++)
            {
                ws.Cells[startRow + i, 13].Value =
                    selectedSample.Efficiencies[i];
            }

            ExcelSignatureHelper.TryAddSignature(ws, "H28");

            wb.Save();
        }
        finally
        {
            wb?.Close(false);

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
        }
    }
    private static void Export_Helper(Excel.Application app,string folder,P2Batch batch,P2GasTest gas)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "Helper_Template.xlsm");

        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 Helper_Template.xlsm");
            return;
        }

        string fileName =
            $"Helper_{gas.GasName}.xlsm";

        string savePath = Path.Combine(folder, fileName);

        File.Copy(templatePath, savePath, true);

        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;

        try
        {
            wb = app.Workbooks.Open(savePath);
            ws = wb.Worksheets["濾網半成品工作表"];

            int currentRow = 2;

            foreach (var sample in gas.Samples)
            {
                ws.Cells[currentRow, 1].Value = batch.ProductionDate;
                ws.Cells[currentRow, 2].Value = batch.TestDate;
                ws.Cells[currentRow, 3].Value = batch.WorkOrder;
                ws.Cells[currentRow, 4].Value = batch.Material;

                ws.Cells[currentRow, 11].Value = sample.Weight;
                ws.Cells[currentRow, 12].Value = sample.PressureDrop;

                if (sample.IsSelected)
                {
                    ws.Cells[currentRow, 13].Value = gas.GasName;
                    ws.Cells[currentRow, 14].Value = gas.Concentration;

                    if (sample.Efficiencies.Count > 0)
                        ws.Cells[currentRow, 15].Value =
                            sample.Efficiencies[0];
                }

                currentRow++;
            }

            wb.Save();
        }
        finally
        {
            wb?.Close(false);

            if (ws != null) Marshal.ReleaseComObject(ws);
            if (wb != null) Marshal.ReleaseComObject(wb);
        }
    }
    private static void WriteTextWithSubscriptNumbers(Excel.Range cell,string text)
    {
        if (cell == null)
            return;
        cell.Value = text;
        if (string.IsNullOrWhiteSpace(text))
            return;
        for (int i = 0; i < text.Length; i++)
        {
            if (char.IsDigit(text[i]))
            {
                Excel.Characters ch = cell.Characters[i + 1, 1];
                ch.Font.Subscript = true;
            }
        }
    }
}
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

        // 先選報告資料夾
        using (var fbd = new FolderBrowserDialog())
        {
            if (fbd.ShowDialog() != DialogResult.OK)
                return;

            string folderPath = fbd.SelectedPath;

            foreach (var gas in batch.GasTests)
            {
                string fileName =
                    $"{batch.ReportNo}_{batch.Material}_{batch.TargetGsm}_{batch.WorkOrder}_{gas.GasName} ({DateTime.Parse(batch.TestDate):MMdd}生產).xlsx";

                string savePath = Path.Combine(folderPath, fileName);

                Export_Report(savePath, batch, gas);
            }
        }

        // 再選 Helper 儲存位置（只做一次）
        using (var sfd = new SaveFileDialog())
        {
            sfd.Filter = "Excel Macro (*.xlsm)|*.xlsm";
            sfd.FileName = "Helper.xlsm";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            Page2HelperExporter.Export(sfd.FileName, batch);
        }
    }
    private static void Export_Report(string savePath, P2Batch batch, P2GasTest gas)
    {
        string templatePath = Path.Combine(
            Application.StartupPath,
            "QC_SemiFinished_Template.xlsx"
        );
        if (!File.Exists(templatePath))
        {
            MessageBox.Show("找不到 QC_SemiFinished_Template.xlsx");
            return;
        }
        foreach (var gasname in batch.GasTests)
        {
            var selectedSample = gas.Samples.FirstOrDefault(s => s.IsSelected);
            if (selectedSample == null) continue;
            var weights = gas.Samples.Select(s => s.Weight ?? 0).ToList();
            var drops = gas.Samples.Select(s => s.PressureDrop ?? 0).ToList();
            File.Copy(templatePath, savePath, true);
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                wb = app.Workbooks.Open(savePath);
                ws = (Excel.Worksheet)wb.Sheets["濾網半成品報告"];
                int idx = gas.Samples.IndexOf(selectedSample);
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
                if (selectedSample != null && selectedSample.Efficiencies != null)
                {
                    for (int i = 0; i < selectedSample.Efficiencies.Count; i++)
                    {
                        ws.Cells[startRow + i, 13].Value = selectedSample.Efficiencies[i];
                    }
                }
                System.Threading.Thread.Sleep(100);
                ExcelSignatureHelper.TryAddSignature(ws, "H28");
                wb.Save();
            }
            finally
            {
                wb?.Close(false);
                app?.Quit();
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) Marshal.ReleaseComObject(app);
            }
        }
    }
    public static class Page2HelperExporter
    {
        public static void Export(string helperSavePath, P2Batch batch)
        {
            if (batch == null || batch.GasTests.Count == 0)
                return;
            string templatePath = Path.Combine(
                Application.StartupPath,
                "Helper_Template.xlsm"
            );
            if (!File.Exists(templatePath))
            {
                MessageBox.Show("找不到 Helper_Template.xlsm");
                return;
            }
            File.Copy(templatePath, helperSavePath, true);
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                wb = app.Workbooks.Open(helperSavePath);
                ws = (Excel.Worksheet)wb.Worksheets["濾網半成品工作表"];
                int startRow =
                    ws.Cells[ws.Rows.Count, 1]
                      .End(Excel.XlDirection.xlUp)
                      .Row + 1;
                int currentRow = startRow;
                DateTime prodDt = DateTime.Parse(batch.ProductionDate);
                foreach (var gas in batch.GasTests)
                {
                    foreach (var sample in gas.Samples)
                    {
                        ws.Cells[currentRow, 1].Value = batch.ProductionDate;
                        ws.Cells[currentRow, 2].Value = batch.TestDate;
                        ws.Cells[currentRow, 3].Value = batch.WorkOrder;
                        ws.Cells[currentRow, 4].Value = batch.Material;
                        ws.Cells[currentRow, 6].Value = batch.TargetGsm;
                        ws.Cells[currentRow, 7].Value = batch.Glue;
                        ws.Cells[currentRow, 8].Value = batch.Speed;
                        ws.Cells[currentRow, 9].Value = batch.Pressure;
                        ws.Cells[currentRow, 10].Value = batch.WindSpeed;
                        ws.Cells[currentRow, 11].Value = sample.Weight;
                        ws.Cells[currentRow, 12].Value = sample.PressureDrop;
                        if (sample.IsSelected)
                        {
                            ws.Cells[currentRow, 13].Value = gas.GasName;
                            ws.Cells[currentRow, 14].Value = gas.Concentration;
                            if (sample.Efficiencies.Count > 0)
                                ws.Cells[currentRow, 15].Value = sample.Efficiencies[0];
                            if (sample.Efficiencies.Count > 10)
                                ws.Cells[currentRow, 16].Value = sample.Efficiencies[10];
                            for (int i = 0; i < sample.Efficiencies.Count; i++)
                            {
                                ws.Cells[3 + i, 21].Value = sample.Efficiencies[i];
                            }
                        }
                        ws.Cells[1, 21].value = $"{batch.TestDate:MM.dd} {batch.Material} {sample.Weight}gsm ({sample.PressureDrop}Pa)-{prodDt:MMdd}生產";
                        currentRow++;
                    }
                }
                wb.Save();
            }
            finally
            {
                wb?.Close(false);
                app?.Quit();

                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (app != null) Marshal.ReleaseComObject(app);
            }
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
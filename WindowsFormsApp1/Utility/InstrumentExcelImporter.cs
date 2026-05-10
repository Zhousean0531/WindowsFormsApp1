using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access;

public static class InstrumentExcelImporter
{
    private static readonly HashSet<string> TargetInstrumentNos =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "QAD-API-01",
            "QAD-API-05",
            "QAD-GT-02",
            "QAD-IAQ-05",
            "QAD-MIT-02",
            "QAD-MIT-03",
            "QAD-PID-01",
            "QAD-GT-03"
        };

    public static void ImportOnStartup()
    {
        string userName = Environment.UserName;

        string excelPath = Path.Combine(
            @"C:\Users",
            userName,
            @"OneDrive - 鈺祥企業股份有限公司",
            @"General - 鈺祥-品保實驗室",
            @"儀器相關資訊",
            @"2.量規儀器管理程序",
            "量規儀器一覽表(含校正)新.xlsx"
        );

        if (!File.Exists(excelPath))
        {
            MessageBox.Show("找不到儀器清單 Excel：" + excelPath);
            return;
        }

        ImportFromExcel(excelPath);
    }

    public static void ImportFromExcel(string excelPath)
    {
        int importCount = 0;

        using (XLWorkbook wb = new XLWorkbook(excelPath))
        {
            IXLWorksheet ws = wb.Worksheet(2);

            int row = 2; // 第 1 列是標題，所以從第 2 列開始

            while (true)
            {
                string instrumentNo = ws.Cell(row, "C").GetString().Trim();

                // C、D、J、K 都空，視為資料結束
                if (string.IsNullOrWhiteSpace(instrumentNo) &&
                    ws.Cell(row, "D").IsEmpty() &&
                    ws.Cell(row, "J").IsEmpty() &&
                    ws.Cell(row, "K").IsEmpty())
                {
                    break;
                }

                if (!TargetInstrumentNos.Contains(instrumentNo))
                {
                    row++;
                    continue;
                }

                string instrumentName = ws.Cell(row, "E").GetString().Trim();

                if (string.IsNullOrWhiteSpace(instrumentName))
                {
                    row++;
                    continue;
                }

                DateTime calibrationDate;
                DateTime expireDate;

                if (!TryGetDate(ws.Cell(row, "J"), out calibrationDate))
                {
                    row++;
                    continue;
                }

                if (!TryGetDate(ws.Cell(row, "K"), out expireDate))
                {
                    row++;
                    continue;
                }

                InstrumentRepository.UpsertByInstrumentName(
                    instrumentName,
                    calibrationDate,
                    expireDate,
                    "SYSTEM"
                );

                importCount++;
                row++;
            }
        }

    }

    private static bool TryGetDate(IXLCell cell, out DateTime date)
    {
        date = DateTime.MinValue;

        if (cell == null || cell.IsEmpty())
            return false;

        if (cell.DataType == XLDataType.DateTime)
        {
            date = cell.GetDateTime();
            return true;
        }

        string text = cell.GetString().Trim();

        if (DateTime.TryParse(text, out date))
            return true;

        return false;
    }
}
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;
public static class MaterialMasterHelper
{
    private static Dictionary<string, MaterialInfo> _map;
    private static string JsonPath =>
        Path.Combine(Application.StartupPath, "MaterialMaster.json");

    private static string CsvPath =>
        Path.Combine(Application.StartupPath, "MaterialList.csv");

    // ===== 主載入 =====
    public static void Load()
    {
        _map = new Dictionary<string, MaterialInfo>();

        LoadFromJson();   // 原本的功能
        LoadFromCsv();    // 新增的功能
    }

    // ===== JSON 讀取
    private static void LoadFromJson()
    {
        if (!File.Exists(JsonPath))
        {
            MessageBox.Show("找不到 MaterialMaster.json");
            return;
        }

        try
        {
            var json = File.ReadAllText(JsonPath);
            _map = JsonConvert.DeserializeObject<Dictionary<string, MaterialInfo>>(json)
                   ?? new Dictionary<string, MaterialInfo>();
        }
        catch
        {
            MessageBox.Show("MaterialMaster.json 讀取失敗");
            _map = new Dictionary<string, MaterialInfo>();
        }
    }
    public static void Add(MaterialInfo info)
    {
        if (_map == null)
            Load();

        if (_map.ContainsKey(info.MaterialNo))
            return;

        _map.Add(info.MaterialNo, info);

        string line = string.Join(",",
        CsvEscape(info.MaterialNo),
        CsvEscape(info.MaterialName),
        CsvEscape(info.InUnit),
        CsvEscape(info.SampleQty),
        CsvEscape(info.InspectUnit),
        CsvEscape(info.Spec)
        );

        File.AppendAllText(CsvPath, Environment.NewLine + line);
    }
    // ===== CSV 讀取
    private static void LoadFromCsv()
    {
        if (!File.Exists(CsvPath))
            return;

        try
        {
            using (var parser = new TextFieldParser(CsvPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                // 跳過 header
                if (!parser.EndOfData)
                    parser.ReadFields();

                while (!parser.EndOfData)
                {
                    var c = parser.ReadFields();
                    if (c == null || c.Length < 6) continue;

                    string materialNo = c[0]
                        .Trim()
                        .Trim('\uFEFF')
                        .ToUpper();

                    if (_map.ContainsKey(materialNo)) continue;

                    _map.Add(materialNo, new MaterialInfo
                    {
                        MaterialNo = materialNo,
                        MaterialName = c[1].Trim(),
                        InUnit = c[2].Trim(),
                        SampleQty = c[3].Trim(),
                        InspectUnit = c[4].Trim(),
                        Spec = c[5]   
                    });
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("MaterialList.csv 讀取失敗\n" + ex.Message);
        }
    }
    private static string CsvEscape(string value)
    {
        if (value.Contains(",") || value.Contains("\n") || value.Contains("\""))
        {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    // ===== 查詢 =====
    public static MaterialInfo Get(string materialNo)
    {
        if (_map == null)
            Load();

        return _map.TryGetValue(materialNo, out var info)
            ? info
            : null;
    }
}

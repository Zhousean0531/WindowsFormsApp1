using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

public static class MaterialMasterHelper
{
    private static Dictionary<string, MaterialInfo> _map;

    public static void Load()
    {
        string path = Path.Combine(
            Application.StartupPath,
            "MaterialMaster.json"
        );

        if (!File.Exists(path))
        {
            MessageBox.Show("找不到 MaterialMaster.json");
            _map = new Dictionary<string, MaterialInfo>();
            return;
        }

        var json = File.ReadAllText(path);

        _map = JsonConvert.DeserializeObject<Dictionary<string, MaterialInfo>>(json)
               ?? new Dictionary<string, MaterialInfo>();
    }

    public static MaterialInfo Get(string material)
    {
        if (_map == null)
            Load(); // 保底

        return _map.TryGetValue(material, out var info) ? info : null;
    }
}

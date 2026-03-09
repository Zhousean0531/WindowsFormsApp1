using System;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

public class AppConfig
{
    public string DBPath { get; set; }
}

public static class ConfigHelper
{
    private static AppConfig _config;

    public static string GetDBPath()
    {
        if (_config != null)
            return _config.DBPath;

        try
        {
            string configPath = Path.Combine(
                Application.StartupPath,
                "config.json");

            if (!File.Exists(configPath))
            {
                MessageBox.Show("找不到 config.json");
                return null;
            }

            string json = File.ReadAllText(configPath);

            _config = JsonConvert.DeserializeObject<AppConfig>(json);

            return _config.DBPath;
        }
        catch (Exception ex)
        {
            MessageBox.Show("讀取設定檔失敗\n" + ex.Message);
            return null;
        }
    }
}
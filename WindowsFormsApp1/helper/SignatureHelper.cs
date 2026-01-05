using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

public static class SignatureHelper
{
    public static string FindSignaturePath()
    {
        string sigDir = Path.Combine(Application.StartupPath, "Signatures");
        if (!Directory.Exists(sigDir))
            return null;

        string username = Environment.UserName;

        // 找所有 username.xxx 的檔案
        var files = Directory.GetFiles(sigDir, username + ".*");

        // 只保留圖片
        var image = files.FirstOrDefault(f =>
        {
            string ext = Path.GetExtension(f).ToLower();
            return ext == ".jpg" || ext == ".jpeg" || ext == ".png";
        });

        return image; // 找不到就回傳 null
    }
}

using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
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
    public static class ExcelSignatureHelper
    {
        public static void TryAddSignature(
    Excel.Worksheet ws,
    string cellAddress
)
        {
            string sigPath = SignatureHelper.FindSignaturePath();

            if (!string.IsNullOrEmpty(sigPath) && File.Exists(sigPath))
            {
                Excel.Range target = ws.Range[cellAddress];

                var pic = ws.Shapes.AddPicture(
                    sigPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)target.Left,
                    (float)target.Top,
                    -1,
                    -1
                );

                pic.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                pic.Width = 110;
                if (cellAddress.Equals("L18", StringComparison.OrdinalIgnoreCase))
                {
                    pic.Width = 160; // Page6 要比較大的簽名
                }
            }
            else
            {
                MessageBox.Show(
                    $"找不到使用者 {Environment.UserName} 的簽名檔，將略過簽名。\n\n" +
                    $"請將簽名圖片放在 Signatures 資料夾，\n" +
                    $"並命名為 {Environment.UserName}.jpg / .png / .jpeg",
                    "簽名檔不存在",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }
        }
    }
}

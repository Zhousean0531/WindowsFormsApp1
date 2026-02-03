using System;
using System.IO;
using System.Windows.Forms;

public static class MasterTableHelper
{
    public static void CopyToOneDrive(string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            MessageBox.Show("找不到已完成的總表.xlsx");
            return;
        }

        string userName = Environment.UserName;

        string targetDir =
            $@"C:\Users\{userName}\OneDrive - 鈺祥企業股份有限公司\General - 鈺祥-品保實驗室";

        if (!Directory.Exists(targetDir))
        {
            MessageBox.Show(
                $"找不到 OneDrive 資料夾：\n{targetDir}"
            );
            return;
        }

        string targetPath = Path.Combine(targetDir, "總表.xlsx");

        // 覆蓋複製（因為你是「發佈最新版」）
        File.Copy(sourcePath, targetPath, true);
    }
}

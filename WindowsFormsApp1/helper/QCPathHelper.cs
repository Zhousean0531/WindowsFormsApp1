using System;
using System.IO;

public static class QCPathHelper
{
    public static string Root =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "QCReport"
        );

    public static string Data => Path.Combine(Root, "data");
    public static string Temp => Path.Combine(Root, "temp");
    public static string Log => Path.Combine(Root, "log");

    public static void Ensure()
    {
        Directory.CreateDirectory(Root);
        Directory.CreateDirectory(Data);
        Directory.CreateDirectory(Temp);
        Directory.CreateDirectory(Log);
    }
}

using System;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelInteropHelper
{
    private const int MaxRetryCount = 25;
    private const int RetryDelayMs = 200;

    private const int ExcelBusyHResult = unchecked((int)0x800AC472);
    private const int RpcCallRejectedHResult = unchecked((int)0x80010001);
    private const int RpcRetryLaterHResult = unchecked((int)0x8001010A);

    public static Excel.Application CreateApplication()
    {
        Excel.Application app = new Excel.Application();
        ConfigureApplication(app);
        return app;
    }

    public static void ConfigureApplication(Excel.Application app)
    {
        if (app == null)
            return;

        Retry(() => { app.Visible = false; });
        Retry(() => { app.DisplayAlerts = false; });
        Retry(() => { app.EnableEvents = false; });
        Retry(() => { app.AskToUpdateLinks = false; });
        Retry(() => { app.ScreenUpdating = false; });

        TrySet(() =>
        {
            app.AutomationSecurity =
                Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        });
    }

    public static Excel.Workbook OpenWorkbook(
        Excel.Application app,
        string path,
        bool readOnly = false
    )
    {
        return Retry(() => app.Workbooks.Open(
            path,
            UpdateLinks: 0,
            ReadOnly: readOnly,
            IgnoreReadOnlyRecommended: true,
            Notify: false,
            AddToMru: false
        ));
    }

    public static void Save(Excel.Workbook wb)
    {
        if (wb == null)
            return;

        Retry(() => { wb.Save(); });
    }

    public static void SaveAs(Excel.Workbook wb, string path)
    {
        if (wb == null)
            return;

        Retry(() => { wb.SaveAs(path); });
    }

    public static void CloseWorkbook(Excel.Workbook wb, bool saveChanges)
    {
        if (wb == null)
            return;

        Retry(() => { wb.Close(saveChanges); });
    }

    public static void Quit(Excel.Application app)
    {
        if (app == null)
            return;

        TrySet(() => { app.DisplayAlerts = false; });
        TrySet(() => { app.EnableEvents = false; });
        TrySet(() => { app.ScreenUpdating = true; });
        Retry(() => { app.Quit(); });
    }

    public static void Retry(Action action)
    {
        Retry<object>(() =>
        {
            action();
            return null;
        });
    }

    public static T Retry<T>(Func<T> action)
    {
        COMException lastException = null;

        for (int i = 0; i < MaxRetryCount; i++)
        {
            try
            {
                return action();
            }
            catch (COMException ex) when (IsRetryable(ex))
            {
                lastException = ex;
                Thread.Sleep(RetryDelayMs);
            }
        }

        throw lastException;
    }

    public static bool IsRetryable(COMException ex)
    {
        if (ex == null)
            return false;

        return ex.HResult == ExcelBusyHResult ||
               ex.HResult == RpcCallRejectedHResult ||
               ex.HResult == RpcRetryLaterHResult;
    }

    private static void TrySet(Action action)
    {
        try
        {
            Retry(action);
        }
        catch (COMException)
        {
        }
    }
}

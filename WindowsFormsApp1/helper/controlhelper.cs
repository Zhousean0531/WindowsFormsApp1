using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class ControlHelper
{
    public static T Find<T>(Control root, string name) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T t && c.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                return t;

            var found = Find<T>(c, name);
            if (found != null)
                return found;
        }
        return null;
    }

    // ⭐ 找全部指定型別
    public static List<T> FindAll<T>(Control root) where T : Control
    {
        return Walk(root).OfType<T>().ToList();
    }

    // 取文字（TextBox / ComboBox / DateTimePicker）
    public static string GetText(Control root, string nameContains)
    {
        foreach (Control c in Walk(root))
        {
            if (c.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            if (c is TextBox tb) return tb.Text.Trim();
            if (c is ComboBox cb) return cb.Text.Trim();
            if (c is DateTimePicker dp) return dp.Value.ToString("yyyy.MM.dd");
        }
        return "";
    }

    // 遞迴走訪
    public static IEnumerable<Control> GetAllControls(Control root)
    {
        foreach (Control c in root.Controls)
        {
            yield return c;

            foreach (var child in GetAllControls(c))
                yield return child;
        }
    }
    private static IEnumerable<Control> Walk(Control root)
    {
        foreach (Control c in root.Controls)
        {
            yield return c;

            foreach (var child in Walk(c))
                yield return child;
        }
    }
}

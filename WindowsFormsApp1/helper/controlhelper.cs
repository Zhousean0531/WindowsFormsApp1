using System;

public static class ControlHelper
{
    public static T FindCtrl<T>(Control root, string nameContains) where T : Control
    {
        foreach (Control c in root.Controls)
        {
            if (c is T t &&
                c.Name.IndexOf(nameContains, StringComparison.OrdinalIgnoreCase) >= 0)
                return t;

            var child = FindCtrl<T>(c, nameContains);
            if (child != null) return child;
        }
        return null;
    }
}

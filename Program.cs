using System;
using System.Windows.Forms;
using WindowsFormsApp1;

public class TabPageButtonExample : Form
{

    [STAThread]
    public static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new mainform1());
    }
}

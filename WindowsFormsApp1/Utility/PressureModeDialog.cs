using System.Windows.Forms;
using static Page3ReportExporter;

public static class PressureModeDialog
{
    public static DialogResult Show(out Page3PressureMode mode)
    {
        mode = Page3PressureMode.Set;

        using (Form form = new Form())
        {
            form.Text = "壓損資料類型";
            form.Width = 300;
            form.Height = 160;
            form.StartPosition = FormStartPosition.CenterParent;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;

            Label label = new Label();
            label.Text = "請選擇壓損資料類型：";
            label.Left = 25;
            label.Top = 20;
            label.Width = 230;

            Button btnSet = new Button();
            btnSet.Text = "整組";
            btnSet.Left = 25;
            btnSet.Top = 65;
            btnSet.Width = 70;
            btnSet.DialogResult = DialogResult.Yes;

            Button btnSingle = new Button();
            btnSingle.Text = "單片";
            btnSingle.Left = 105;
            btnSingle.Top = 65;
            btnSingle.Width = 70;
            btnSingle.DialogResult = DialogResult.No;

            Button btnCancel = new Button();
            btnCancel.Text = "取消";
            btnCancel.Left = 185;
            btnCancel.Top = 65;
            btnCancel.Width = 70;
            btnCancel.DialogResult = DialogResult.Cancel;

            form.Controls.Add(label);
            form.Controls.Add(btnSet);
            form.Controls.Add(btnSingle);
            form.Controls.Add(btnCancel);

            form.AcceptButton = btnSet;
            form.CancelButton = btnCancel;

            DialogResult result = form.ShowDialog();

            if (result == DialogResult.Yes)
            {
                mode = Page3PressureMode.Set;
            }
            else if (result == DialogResult.No)
            {
                mode = Page3PressureMode.Single;
            }

            return result;
        }
    }
}
using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private void InitializeClearCurrentPageButton()
        {
            if (ClearCurrentPageButton == null)
                return;

            ClearCurrentPageButton.Click -= ClearCurrentPageButton_Click;
            ClearCurrentPageButton.Click += ClearCurrentPageButton_Click;
            ClearCurrentPageButton.BringToFront();
        }

        private void ClearCurrentPageButton_Click(object sender, EventArgs e)
        {
            ClearCurrentTabData();
            MessageBox.Show("已清空目前分頁資料，日期欄位保留。");
        }

        private void ClearCurrentTabData()
        {
            var page = tabControl1?.SelectedTab;
            if (page == null)
                return;

            foreach (Control control in page.Controls)
                ClearControlValue(control);

            if (page == QueryPage)
            {
                _queryResultTable = null;
                if (QueryExportButton != null)
                    QueryExportButton.Enabled = false;

                SetQueryConditionState();
            }
        }

        private void ClearControlValue(Control control)
        {
            if (control == null || control is DateTimePicker)
                return;

            if (control is TextBox textBox)
            {
                textBox.Clear();
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.Checked = false;
            }
            else if (control is ComboBox comboBox)
            {
                comboBox.SelectedIndex = -1;
                comboBox.Text = string.Empty;
            }
            else if (control is DataGridView grid)
            {
                ClearDataGridView(grid);
            }

            foreach (Control child in control.Controls)
                ClearControlValue(child);
        }

        private void ClearDataGridView(DataGridView grid)
        {
            if (grid == null)
                return;

            try
            {
                grid.EndEdit();
            }
            catch
            {
            }

            grid.DataSource = null;

            try
            {
                grid.Rows.Clear();
            }
            catch (InvalidOperationException)
            {
                for (int i = grid.Rows.Count - 1; i >= 0; i--)
                {
                    if (!grid.Rows[i].IsNewRow)
                        grid.Rows.RemoveAt(i);
                }
            }
        }
    }
}

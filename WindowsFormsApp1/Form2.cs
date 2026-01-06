using System;
using System.Collections.Generic;
using System.Windows.Forms;
//選擇檢驗樣品表單
namespace YourNamespace   
{
    public partial class Form2 : Form
    {
        public int SelectedIndex0 { get; private set; } = -1;

        public Form2()
        {
            InitializeComponent();
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        // 另外一個建構式（可選）
        public Form2(List<double> values, string caption = "請選擇第幾筆") : this()
        {
            this.Text = caption;
            var items = new List<string>();
            for (int i = 0; i < values.Count; i++)
                items.Add($"{values[i]}");
            comboBox1.DataSource = items;
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex >= 0)
            {
                SelectedIndex0 = comboBox1.SelectedIndex;
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("請先選擇一筆資料！");
            }
        }
    }
}

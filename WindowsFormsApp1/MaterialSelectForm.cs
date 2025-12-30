using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class MaterialSelectForm : Form
    {
        public List<string> SelectedItems { get; private set; } = new List<string>();

        public MaterialSelectForm(string[] items)
        {
            InitializeComponent(); // ⬅ 這行才不會報錯！

            // 以下是動態建立介面
            FlowLayoutPanel panel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            foreach (var item in items)
            {
                var cb = new CheckBox
                {
                    Text = item,
                    AutoSize = true,
                    Checked = true
                };
                panel.Controls.Add(cb);
            }

            var btn = new Button
            {
                Text = "確定",
                Size = new Size(100, 40),
                Dock = DockStyle.Bottom,
                Margin = new Padding(10)
            };


            btn.Click += (s, e) =>
            {
                foreach (var ctrl in panel.Controls)
                {
                    if (ctrl is CheckBox cb && cb.Checked)
                        SelectedItems.Add(cb.Text);
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.Add(panel);
            this.Controls.Add(btn);
            this.Text = "請選擇測試批號";
            this.Size = new System.Drawing.Size(300, 300);
        }
    }

}

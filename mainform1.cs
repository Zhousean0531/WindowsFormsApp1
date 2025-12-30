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
    public partial class mainform1 : Form
    {
        public mainform1()
        {
            InitializeComponent();
            this.Load += new EventHandler(this.mainform1_Load); // 綁定 Load 事件
        }

        private void mainform1_Load(object sender, EventArgs e)
        {
            MessageBox.Show(textBox36.Parent.Name);  // 顯示父容器名稱
        }
    }
}

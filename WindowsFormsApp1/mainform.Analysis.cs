using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private ToolTip QualityAnalysisToolTip;

        private void InitializeQualityAnalysisLauncher()
        {
            if (QualityAnalysisImageButton == null)
                return;

            QualityAnalysisImageButton.Image = BuildQualityAnalysisIcon(QualityAnalysisImageButton.Size);

            QualityAnalysisToolTip = new ToolTip();
            QualityAnalysisToolTip.SetToolTip(QualityAnalysisImageButton, "品保數據分析");

            QualityAnalysisImageButton.Click -= QualityAnalysisImageButton_Click;
            QualityAnalysisImageButton.Click += QualityAnalysisImageButton_Click;
            QualityAnalysisImageButton.BringToFront();
        }

        private void QualityAnalysisImageButton_Click(object sender, EventArgs e)
        {
            using (var form = new QualityAnalysisForm())
            {
                form.ShowDialog(this);
            }
        }

        private Image BuildQualityAnalysisIcon(Size size)
        {
            var bmp = new Bitmap(size.Width, size.Height);

            using (var g = Graphics.FromImage(bmp))
            using (var borderPen = new Pen(Color.FromArgb(185, 185, 185), 1F))
            using (var axisPen = new Pen(Color.FromArgb(105, 105, 105), 1.4F))
            using (var uslPen = new Pen(Color.Red, 2F))
            using (var avgPen = new Pen(Color.FromArgb(65, 155, 55), 2F))
            using (var dotBrush = new SolidBrush(Color.FromArgb(20, 158, 208)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.White);
                g.DrawRectangle(borderPen, 1, 1, size.Width - 3, size.Height - 3);

                var plot = new Rectangle(9, 10, size.Width - 18, size.Height - 22);
                g.DrawLine(axisPen, plot.Left, plot.Top, plot.Left, plot.Bottom);
                g.DrawLine(axisPen, plot.Left, plot.Bottom, plot.Right, plot.Bottom);

                g.DrawLine(uslPen, plot.Left + 1, plot.Top + 7, plot.Right - 2, plot.Top + 7);
                g.DrawLine(uslPen, plot.Left + 1, plot.Bottom - 6, plot.Right - 2, plot.Bottom - 6);
                g.DrawLine(avgPen, plot.Left + 1, plot.Top + plot.Height / 2, plot.Right - 2, plot.Top + plot.Height / 2);

                for (int i = 0; i < 17; i++)
                {
                    int x = plot.Left + 4 + i * 2;
                    int y = plot.Top + 9 + ((i * 7) % 18);
                    g.FillEllipse(dotBrush, x, y, 3, 3);
                }
            }

            return bmp;
        }
    }
}

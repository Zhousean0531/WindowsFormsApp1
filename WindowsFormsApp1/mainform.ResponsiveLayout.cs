using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private const int HeaderHeight = 50;
        private const int HeaderButtonTop = 12;
        private const int HeaderRightMargin = 31;
        private const int HeaderButtonGap = 37;
        private const int AnalysisButtonGap = 17;
        private const int PageScrollMargin = 24;

        private void InitializeResponsiveLayout()
        {
            SuspendLayout();

            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimumSize = new Size(900, 650);

            tabControl1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            SearchButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            execute.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            if (QualityAnalysisImageButton != null)
                QualityAnalysisImageButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            foreach (TabPage page in tabControl1.TabPages)
                ConfigureScrollablePage(page);

            Resize += (s, e) => UpdateResponsiveLayout();
            tabControl1.ControlAdded += (s, e) =>
            {
                if (e.Control is TabPage page)
                    ConfigureScrollablePage(page);
            };

            UpdateResponsiveLayout();
            ResumeLayout(false);
        }

        private void UpdateResponsiveLayout()
        {
            if (tabControl1 == null)
                return;

            tabControl1.SetBounds(
                0,
                HeaderHeight,
                Math.Max(320, ClientSize.Width),
                Math.Max(220, ClientSize.Height - HeaderHeight));

            PositionHeaderButtons();

            foreach (TabPage page in tabControl1.TabPages)
                UpdatePageScrollArea(page);
        }

        private void PositionHeaderButtons()
        {
            if (execute == null || SearchButton == null)
                return;

            execute.Top = HeaderButtonTop;
            SearchButton.Top = HeaderButtonTop;

            execute.Left = Math.Max(
                HeaderRightMargin,
                ClientSize.Width - execute.Width - HeaderRightMargin);

            SearchButton.Left = Math.Max(
                HeaderRightMargin,
                execute.Left - SearchButton.Width - HeaderButtonGap);

            if (QualityAnalysisImageButton != null)
            {
                QualityAnalysisImageButton.Top = HeaderButtonTop;
                QualityAnalysisImageButton.Left = Math.Max(
                    HeaderRightMargin,
                    SearchButton.Left - QualityAnalysisImageButton.Width - AnalysisButtonGap);
            }
        }

        private void ConfigureScrollablePage(TabPage page)
        {
            if (page == null)
                return;

            page.AutoScroll = true;
            page.AutoScrollMargin = new Size(PageScrollMargin, PageScrollMargin);
            page.Resize += (s, e) => UpdatePageScrollArea(page);
            page.ControlAdded += (s, e) => UpdatePageScrollArea(page);
            page.ControlRemoved += (s, e) => UpdatePageScrollArea(page);
            UpdatePageScrollArea(page);
        }

        private void UpdatePageScrollArea(TabPage page)
        {
            if (page == null || page.Controls.Count == 0)
                return;

            int right = 0;
            int bottom = 0;

            foreach (Control control in page.Controls.Cast<Control>().Where(x => x.Visible))
            {
                right = Math.Max(right, control.Right + PageScrollMargin);
                bottom = Math.Max(bottom, control.Bottom + PageScrollMargin);
            }

            page.AutoScrollMinSize = new Size(right, bottom);
        }
    }
}

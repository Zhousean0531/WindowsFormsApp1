using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1
    {
        private const int HeaderHeight = 74;
        private const int HeaderButtonTop = 8;
        private const int HeaderRightMargin = 24;
        private const int HeaderButtonGap = 16;
        private const int AnalysisButtonGap = 12;
        private static readonly Size HeaderButtonSize = new Size(100, 42);
        private static readonly Size ClearButtonSize = new Size(92, 42);
        private static readonly Size AnalysisButtonSize = new Size(42, 42);
        private const int PageScrollMargin = 24;
        private readonly Dictionary<Control, Rectangle> responsiveBaseBounds =
            new Dictionary<Control, Rectangle>();

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
            SearchButton.Size = HeaderButtonSize;
            execute.Size = HeaderButtonSize;
            SearchButton.Font = new Font("微軟正黑體", 13.5F, FontStyle.Regular, GraphicsUnit.Point, 136);
            execute.Font = new Font("微軟正黑體", 13.5F, FontStyle.Regular, GraphicsUnit.Point, 136);

            if (ClearCurrentPageButton != null)
            {
                ClearCurrentPageButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                ClearCurrentPageButton.Size = ClearButtonSize;
                ClearCurrentPageButton.Font = new Font("微軟正黑體", 13.5F, FontStyle.Regular, GraphicsUnit.Point, 136);
            }

            if (QualityAnalysisImageButton != null)
            {
                QualityAnalysisImageButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                QualityAnalysisImageButton.Size = AnalysisButtonSize;
                QualityAnalysisImageButton.Image = BuildQualityAnalysisIcon(AnalysisButtonSize);
            }

            ConfigureResponsiveControls();

            foreach (TabPage page in tabControl1.TabPages)
                ConfigureScrollablePage(page);

            Resize += (s, e) => UpdateResponsiveLayout();
            tabControl1.ControlAdded += (s, e) =>
            {
                if (e.Control is TabPage page)
                {
                    ConfigureScrollablePage(page);
                    ConfigureResponsivePage(page);
                }
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
            {
                ConfigureResponsivePage(page);
                UpdatePageScrollArea(page);
            }
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

            if (ClearCurrentPageButton != null)
            {
                ClearCurrentPageButton.Top = HeaderButtonTop;
                ClearCurrentPageButton.Left = Math.Max(
                    HeaderRightMargin,
                    SearchButton.Left - ClearCurrentPageButton.Width - HeaderButtonGap);
            }

            if (QualityAnalysisImageButton != null)
            {
                QualityAnalysisImageButton.Top = HeaderButtonTop;
                int anchorLeft = ClearCurrentPageButton != null
                    ? ClearCurrentPageButton.Left
                    : SearchButton.Left;

                QualityAnalysisImageButton.Left = Math.Max(
                    HeaderRightMargin,
                    anchorLeft - QualityAnalysisImageButton.Width - AnalysisButtonGap);
            }
        }

        private void ConfigureScrollablePage(TabPage page)
        {
            if (page == null)
                return;

            page.AutoScroll = true;
            page.AutoScrollMargin = new Size(PageScrollMargin, PageScrollMargin);
            page.Resize += (s, e) =>
            {
                ConfigureResponsivePage(page);
                UpdatePageScrollArea(page);
            };
            page.ControlAdded += (s, e) => UpdatePageScrollArea(page);
            page.ControlRemoved += (s, e) => UpdatePageScrollArea(page);
            UpdatePageScrollArea(page);
        }

        private void ConfigureResponsivePage(TabPage page)
        {
            if (page == null)
                return;

            if (page == FilterRawPage || page.Name == "FilterRawPage")
                LayoutPage1();
            else if (page == FilterInProcessPage || page.Name == "FilterInProcessPage")
                LayoutPage2();
            else if (page == FilterPage || page.Name == "FilterPage")
                LayoutPage3();
            else if (page == CylinderRawPage || page.Name == "CylinderRawPage")
                LayoutPage4();
            else if (page == CylinderPage || page.Name == "CylinderPage")
                LayoutPage5();
            else if (page == RawMaterialPage || page.Name == "RawMaterialPage")
                LayoutPage6();
        }

        private void ConfigureResponsiveControls()
        {
            ConfigureResponsiveGrid(FilterRawParticleSizeBox);
            ConfigureResponsiveGrid(FilterBox);
            ConfigureResponsiveGrid(CylinderRawMeshBox);
            ConfigureResponsiveGrid(CylinderBox);
            ConfigureResponsiveGrid(RawMaterialdgv);

            ConfigureResponsiveTable(FilterInProcessEffPanel);
            ConfigureResponsiveTable(CylinderRawEffPanel);

            if (FilterRawEffvalueBox != null)
                FilterRawEffvalueBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        }

        private void ConfigureResponsiveGrid(DataGridView grid)
        {
            if (grid == null)
                return;

            grid.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Bottom |
                AnchorStyles.Left |
                AnchorStyles.Right;

            grid.ScrollBars = ScrollBars.Both;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.AllowUserToResizeColumns = true;
            grid.AllowUserToResizeRows = true;

            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.Resizable = DataGridViewTriState.True;

                if (column.Width < 80)
                    column.Width = 80;
            }
        }

        private void ConfigureResponsiveTable(TableLayoutPanel panel)
        {
            if (panel == null)
                return;

            panel.Anchor =
                AnchorStyles.Top |
                AnchorStyles.Bottom |
                AnchorStyles.Left |
                AnchorStyles.Right;

            foreach (Control control in panel.Controls)
            {
                if (control is Label label)
                {
                    label.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                    label.TextAlign = ContentAlignment.MiddleCenter;
                }
                else
                {
                    control.Anchor =
                        AnchorStyles.Top |
                        AnchorStyles.Bottom |
                        AnchorStyles.Left |
                        AnchorStyles.Right;
                }
            }
        }

        private void LayoutPage1()
        {
            LayoutFromBase(FilterRawPage, FilterRawEffvalueBox, growWidth: true, growHeight: false);
            LayoutFromBase(FilterRawPage, FilterRawParticleSizeBox, growWidth: true, growHeight: true);
        }

        private void LayoutPage2()
        {
            LayoutFromBase(FilterInProcessPage, FilterInProcessEffPanel, growWidth: false, growHeight: false);
        }

        private void LayoutPage3()
        {
            LayoutFromBase(FilterPage, FilterBox, growWidth: true, growHeight: true);
        }

        private void LayoutPage4()
        {
            LayoutFromBase(CylinderRawPage, CylinderRawEffPanel, growWidth: false, growHeight: false);
            LayoutFromBase(CylinderRawPage, CylinderRawMeshBox, growWidth: false, growHeight: false);
        }

        private void LayoutPage5()
        {
            LayoutFromBase(CylinderPage, CylinderBox, growWidth: true, growHeight: true);
        }

        private void LayoutPage6()
        {
            if (RawMaterialPage == null || RawMaterialdgv == null)
                return;

            LayoutFromBase(RawMaterialPage, RawMaterialdgv, growWidth: true, growHeight: true);
        }

        private void LayoutFromBase(TabPage page, Control control, bool growWidth, bool growHeight)
        {
            if (page == null || control == null)
                return;

            Rectangle baseBounds = GetResponsiveBaseBounds(control);

            int width = Math.Max(
                baseBounds.Width,
                page.ClientSize.Width - baseBounds.Left - PageScrollMargin);

            int height = Math.Max(
                baseBounds.Height,
                page.ClientSize.Height - baseBounds.Top - PageScrollMargin);

            if (!growWidth)
                width = baseBounds.Width;

            if (!growHeight)
                height = baseBounds.Height;

            control.Bounds = new Rectangle(baseBounds.Left, baseBounds.Top, width, height);
        }

        private Rectangle GetResponsiveBaseBounds(Control control)
        {
            if (!responsiveBaseBounds.TryGetValue(control, out Rectangle bounds))
            {
                bounds = control.Bounds;
                responsiveBaseBounds[control] = bounds;
            }

            return bounds;
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

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1.Data_Access.Analysis;

namespace WindowsFormsApp1
{
    public class QualityAnalysisForm : Form
    {
        private readonly QualityAnalysisChartControl chart;
        private readonly CheckBox chkPlus1Sigma;
        private readonly CheckBox chkMinus1Sigma;
        private readonly CheckBox chkPlus2Sigma;
        private readonly CheckBox chkMinus2Sigma;
        private readonly CheckBox chkPlus3Sigma;
        private readonly CheckBox chkMinus3Sigma;
        private readonly CheckBox chkAverage;
        private readonly CheckBox chkUsl;
        private readonly CheckBox chkLsl;
        private readonly DateTimePicker startDateBox;
        private readonly DateTimePicker endDateBox;
        private readonly TextBox materialBox;
        private readonly ComboBox metricBox;
        private readonly Button queryButton;
        private readonly Label statusLabel;

        public QualityAnalysisForm()
        {
            Text = "QC Report 品保數據分析";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(1120, 720);
            Size = new Size(1180, 760);
            BackColor = Color.White;
            Font = new Font("Microsoft JhengHei UI", 10F, FontStyle.Regular);

            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 136,
                BackColor = Color.White,
                Padding = new Padding(24, 16, 24, 12)
            };

            var title = new Label
            {
                Text = "品保數據分析",
                Font = new Font("Microsoft JhengHei UI", 18F, FontStyle.Bold),
                Location = new Point(24, 18),
                AutoSize = true
            };

            startDateBox = BuildDateBox(24, 72);
            startDateBox.Value = DateTime.Today.AddMonths(-6);
            endDateBox = BuildDateBox(216, 72);
            materialBox = BuildTextBox("IKP201", 452, 72, 150);
            metricBox = BuildComboBox(680, 72, 150);
            queryButton = BuildCommandButton("查詢", 870, 68, true);
            queryButton.Click += QueryButton_Click;
            statusLabel = new Label
            {
                Text = "請輸入條件後按查詢",
                Font = new Font("Microsoft JhengHei UI", 9F),
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(24, 108),
                AutoSize = true
            };

            topPanel.Controls.Add(title);
            topPanel.Controls.Add(BuildLabel("查詢期間", 24, 50));
            topPanel.Controls.Add(startDateBox);
            topPanel.Controls.Add(BuildLabel("~", 196, 76));
            topPanel.Controls.Add(endDateBox);
            topPanel.Controls.Add(BuildLabel("料號 / 種類", 452, 50));
            topPanel.Controls.Add(materialBox);
            topPanel.Controls.Add(BuildLabel("分析項目", 680, 50));
            topPanel.Controls.Add(metricBox);
            topPanel.Controls.Add(queryButton);
            topPanel.Controls.Add(statusLabel);

            chart = new QualityAnalysisChartControl
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };
            chart.PointLegendText = GetPointLegendText(metricBox.Text);
            chart.MetricKey = GetMetricKey(metricBox.Text);

            var rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 230,
                BackColor = Color.FromArgb(248, 248, 248),
                Padding = new Padding(18, 18, 18, 18)
            };

            var panelTitle = new Label
            {
                Text = "顯示線條",
                Font = new Font("Microsoft JhengHei UI", 13F, FontStyle.Bold),
                Location = new Point(18, 18),
                AutoSize = true
            };

            chkPlus1Sigma = BuildLineCheckBox("+1σ", 18, 62, true);
            chkMinus1Sigma = BuildLineCheckBox("-1σ", 18, 94, true);
            chkPlus2Sigma = BuildLineCheckBox("+2σ", 18, 126, true);
            chkMinus2Sigma = BuildLineCheckBox("-2σ", 18, 158, true);
            chkPlus3Sigma = BuildLineCheckBox("+3σ", 18, 190, true);
            chkMinus3Sigma = BuildLineCheckBox("-3σ", 18, 222, true);
            chkAverage = BuildLineCheckBox("平均值", 18, 268, true);
            chkUsl = BuildLineCheckBox("USL", 18, 300, true);
            chkLsl = BuildLineCheckBox("LSL", 18, 332, true);

            rightPanel.Controls.Add(panelTitle);
            rightPanel.Controls.Add(chkPlus1Sigma);
            rightPanel.Controls.Add(chkMinus1Sigma);
            rightPanel.Controls.Add(chkPlus2Sigma);
            rightPanel.Controls.Add(chkMinus2Sigma);
            rightPanel.Controls.Add(chkPlus3Sigma);
            rightPanel.Controls.Add(chkMinus3Sigma);
            rightPanel.Controls.Add(chkAverage);
            rightPanel.Controls.Add(chkUsl);
            rightPanel.Controls.Add(chkLsl);

            Controls.Add(chart);
            Controls.Add(rightPanel);
            Controls.Add(topPanel);

            ApplyLineVisibility();
        }

        private DateTimePicker BuildDateBox(int x, int y)
        {
            return new DateTimePicker
            {
                Format = DateTimePickerFormat.Custom,
                CustomFormat = "yyyy.MM.dd",
                Location = new Point(x, y),
                Size = new Size(150, 28),
                Font = Font
            };
        }

        private TextBox BuildTextBox(string text, int x, int y, int width)
        {
            return new TextBox
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 28),
                Font = Font
            };
        }

        private ComboBox BuildComboBox(int x, int y, int width)
        {
            var box = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(x, y),
                Size = new Size(width, 28),
                Font = Font
            };

            box.Items.AddRange(new object[] { "重量", "壓損", "效率(%)", "Particle", "TVOC" });
            box.SelectedIndex = 0;

            return box;
        }

        private string GetPointLegendText(string metricName)
        {
            switch ((metricName ?? "").Trim())
            {
                case "重量":
                    return "重量(Kg)";
                case "Particle":
                    return "Particle(0.5μm)";
                case "TVOC":
                    return "TVOC";
                case "壓損":
                    return "Dressure Drop";
                default:
                    return metricName;
            }
        }

        private string GetMetricKey(string metricName)
        {
            switch ((metricName ?? "").Trim())
            {
                case "重量":
                    return "Weight";
                case "Particle":
                    return "Particle";
                case "TVOC":
                    return "TVOC";
                case "壓損":
                    return "PressureDrop";
                case "效率(%)":
                    return "Efficiency";
                default:
                    return metricName;
            }
        }

        private Button BuildCommandButton(string text, int x, int y, bool enabled)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(88, 34),
                Font = Font,
                Enabled = enabled
            };
        }

        private Label BuildLabel(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y),
                Font = Font,
                AutoSize = true
            };
        }

        private CheckBox BuildLineCheckBox(string text, int x, int y, bool isChecked)
        {
            var box = new CheckBox
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(170, 26),
                Checked = isChecked,
                Font = Font,
                AutoSize = false
            };

            box.CheckedChanged += (s, e) => ApplyLineVisibility();

            return box;
        }

        private void ApplyLineVisibility()
        {
            chart.ShowPlus1Sigma = chkPlus1Sigma.Checked;
            chart.ShowMinus1Sigma = chkMinus1Sigma.Checked;
            chart.ShowPlus2Sigma = chkPlus2Sigma.Checked;
            chart.ShowMinus2Sigma = chkMinus2Sigma.Checked;
            chart.ShowPlus3Sigma = chkPlus3Sigma.Checked;
            chart.ShowMinus3Sigma = chkMinus3Sigma.Checked;
            chart.ShowAverage = chkAverage.Checked;
            chart.ShowUsl = chkUsl.Checked;
            chart.ShowLsl = chkLsl.Checked;
            chart.Invalidate();
        }

        private void QueryButton_Click(object sender, EventArgs e)
        {
            RunAnalysisQuery();
        }

        private void RunAnalysisQuery()
        {
            string material = (materialBox.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(material))
            {
                MessageBox.Show("請輸入料號 / 種類");
                materialBox.Focus();
                return;
            }

            DateTime start = startDateBox.Value.Date;
            DateTime end = endDateBox.Value.Date;

            if (start > end)
            {
                MessageBox.Show("查詢起始日期不可大於結束日期");
                return;
            }

            string metricKey = GetMetricKey(metricBox.Text);

            try
            {
                queryButton.Enabled = false;
                statusLabel.Text = "查詢中...";

                var result = QualityAnalysisRepository.Query(material, metricKey, start, end);

                chart.PointLegendText = GetPointLegendText(metricBox.Text);
                chart.MetricKey = metricKey;
                chart.SetResult(material, result);

                string settingText = result.Setting == null ? "；未找到設定表資料，線條依查詢期間計算" : "";
                if (result.SigmaPoints.Count == 0)
                    settingText += "；標準差期間無資料，平均值與σ線不顯示";

                statusLabel.Text =
                    $"查詢完成：{result.Points.Count} 點；標準差期間 {result.SigmaStartDate:yyyy.MM.dd} ~ {result.SigmaEndDate:yyyy.MM.dd}，{result.SigmaPoints.Count} 點{settingText}";
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢品保分析資料時發生錯誤：\n" + ex.Message);
                statusLabel.Text = "查詢失敗";
            }
            finally
            {
                queryButton.Enabled = true;
            }
        }
    }

    internal class QualityAnalysisChartControl : Control
    {
        private sealed class YAxisRule
        {
            public YAxisRule(double minimum, double maximum, double step, string labelFormat)
            {
                Minimum = minimum;
                Maximum = maximum;
                Step = step;
                LabelFormat = labelFormat;
            }

            public double Minimum { get; }
            public double Maximum { get; }
            public double Step { get; }
            public string LabelFormat { get; }
        }

        private readonly List<QualityAnalysisPoint> points = new List<QualityAnalysisPoint>();
        private string titleText = "IKP201";
        private double? average;
        private double? sigma;
        private double? usl;
        private double? lsl;

        public bool ShowPlus1Sigma { get; set; }
        public bool ShowMinus1Sigma { get; set; }
        public bool ShowPlus2Sigma { get; set; }
        public bool ShowMinus2Sigma { get; set; }
        public bool ShowPlus3Sigma { get; set; }
        public bool ShowMinus3Sigma { get; set; }
        public bool ShowAverage { get; set; }
        public bool ShowUsl { get; set; }
        public bool ShowLsl { get; set; }
        public string PointLegendText { get; set; } = "重量(Kg)";
        public string MetricKey { get; set; } = "Weight";
        public string EmptyMessage { get; set; } = "請輸入條件後按查詢";

        public QualityAnalysisChartControl()
        {
            DoubleBuffered = true;

            var menu = new ContextMenuStrip();
            var copyImageItem = new ToolStripMenuItem("Copy Image");
            copyImageItem.Click += CopyImageItem_Click;
            menu.Items.Add(copyImageItem);
            ContextMenuStrip = menu;
        }

        private void CopyImageItem_Click(object sender, EventArgs e)
        {
            if (Width <= 0 || Height <= 0)
                return;

            using (var bitmap = new Bitmap(Width, Height))
            {
                DrawToBitmap(bitmap, new Rectangle(0, 0, Width, Height));
                Clipboard.SetImage((Image)bitmap.Clone());
            }
        }

        public void SetResult(string title, QualityAnalysisResult result)
        {
            titleText = string.IsNullOrWhiteSpace(title) ? "" : title.Trim();
            points.Clear();

            if (result != null && result.Points != null)
                points.AddRange(result.Points);

            average = result?.Average;
            sigma = result?.Sigma;
            usl = result?.Setting?.USL;
            lsl = result?.Setting?.LSL;
            EmptyMessage = points.Count == 0 ? "查無符合條件的資料" : "";

            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(Color.White);

            using (var titleFont = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold))
            using (var axisFont = new Font("Microsoft JhengHei UI", 12.5F, FontStyle.Regular))
            using (var legendFont = new Font("Microsoft JhengHei UI", 12F, FontStyle.Regular))
            {
                Rectangle plot = new Rectangle(96, 62, Math.Max(260, Width - 160), Math.Max(220, Height - 250));
                double minY;
                double maxY;
                ResolveYAxis(out minY, out maxY);

                DrawTitle(e.Graphics, titleFont, titleText, plot);
                DrawAxis(e.Graphics, axisFont, plot, minY, maxY);

                if (points.Count == 0)
                    DrawEmptyMessage(e.Graphics, axisFont, plot);

                DrawSamplePoints(e.Graphics, plot, minY, maxY);

                using (var redSolid = new Pen(Color.Red, 2.5F))
                using (var redDash = new Pen(Color.Red, 2.5F) { DashStyle = DashStyle.Dash })
                using (var greenSolid = new Pen(Color.FromArgb(64, 160, 46), 2.5F))
                {
                    if (ShowUsl && usl.HasValue) DrawHorizontalLine(e.Graphics, plot, minY, maxY, usl.Value, redSolid);
                    if (ShowLsl && lsl.HasValue) DrawHorizontalLine(e.Graphics, plot, minY, maxY, lsl.Value, redSolid);
                    if (ShowAverage && average.HasValue) DrawHorizontalLine(e.Graphics, plot, minY, maxY, average.Value, greenSolid);

                    if (average.HasValue && sigma.HasValue)
                    {
                        if (ShowPlus1Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(1), redDash);
                        if (ShowMinus1Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(-1), redDash);
                        if (ShowPlus2Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(2), redDash);
                        if (ShowMinus2Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(-2), redDash);
                        if (ShowPlus3Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(3), redDash);
                        if (ShowMinus3Sigma) DrawHorizontalLine(e.Graphics, plot, minY, maxY, GetSigmaLineValue(-3), redDash);
                    }
                }

                DrawDateLabels(e.Graphics, axisFont, plot);
                DrawLegend(e.Graphics, legendFont, plot);
            }
        }

        private void ResolveYAxis(out double minY, out double maxY)
        {
            var rule = GetYAxisRule();
            if (rule != null)
            {
                minY = rule.Minimum;
                maxY = rule.Maximum;
                return;
            }

            var values = points.Select(x => x.Value).ToList();

            if (values.Count == 0)
            {
                minY = 0;
                maxY = 1;
                return;
            }

            double pointMin = values.Min();
            double pointMax = values.Max();

            minY = pointMin < 0 ? Math.Floor(pointMin) : 0;
            maxY = Math.Ceiling(pointMax);

            if (maxY <= minY)
                maxY = minY + 1;
        }

        private void DrawTitle(Graphics g, Font font, string text, Rectangle plot)
        {
            SizeF size = g.MeasureString(text, font);
            g.DrawString(text, font, Brushes.Black, plot.Left + (plot.Width - size.Width) / 2F, 18);
        }

        private void DrawAxis(Graphics g, Font font, Rectangle plot, double minY, double maxY)
        {
            using (var axisPen = new Pen(Color.FromArgb(170, 170, 170), 1F))
            {
                g.DrawLine(axisPen, plot.Left, plot.Top, plot.Left, plot.Bottom);
                g.DrawLine(axisPen, plot.Left, plot.Bottom, plot.Right, plot.Bottom);

                var rule = GetYAxisRule();
                if (rule != null)
                {
                    DrawFixedAxisTicks(g, font, plot, minY, maxY, rule);
                    return;
                }

                int tickCount = 6;
                for (int i = 0; i <= tickCount; i++)
                {
                    double value = minY + (maxY - minY) * i / tickCount;
                    float py = MapY(plot, minY, maxY, value);
                    DrawYAxisLabel(g, font, plot, py, FormatNumber(value));
                }
            }
        }

        private void DrawFixedAxisTicks(
            Graphics g,
            Font font,
            Rectangle plot,
            double minY,
            double maxY,
            YAxisRule rule)
        {
            int guard = 0;
            for (double value = rule.Minimum; value <= rule.Maximum + rule.Step / 2.0 && guard < 1000; value += rule.Step, guard++)
            {
                double tickValue = Math.Round(value, 10);
                float py = MapY(plot, minY, maxY, tickValue);
                DrawYAxisLabel(g, font, plot, py, tickValue.ToString(rule.LabelFormat));
            }
        }

        private void DrawYAxisLabel(Graphics g, Font font, Rectangle plot, float y, string text)
        {
            SizeF size = g.MeasureString(text, font);
            g.DrawString(text, font, Brushes.Black, plot.Left - size.Width - 14, y - size.Height / 2F);
        }

        private YAxisRule GetYAxisRule()
        {
            string material = NormalizeMaterial(titleText);

            if (IsFixedCylinderMaterial(material))
            {
                if (MetricKey == "Weight")
                    return new YAxisRule(2.50, 5.00, 0.50, "0.00");

                if (MetricKey == "Particle")
                {
                    if (IsMaterial(material, "IKP201", "IKP205", "CI001"))
                        return new YAxisRule(0, 4000, 1000, "0");

                    if (IsMaterial(material, "SG017_A", "SG017_C", "SG017_D", "SG029", "SG035"))
                        return new YAxisRule(0, 10000, 1000, "0");
                }

                if (MetricKey == "TVOC")
                    return new YAxisRule(0, 5.0, 0.5, "0.0");

                if (MetricKey == "PressureDrop")
                    return new YAxisRule(0, 150, 10, "0");
            }

            if (IsFixedFilterParticleMaterial(material))
            {
                if (MetricKey == "Particle")
                    return new YAxisRule(0, 10000, 1000, "0");
            }

            if (MetricKey == "TVOC" && IsFixedFilterTvocMaterial(material))
                return new YAxisRule(0, 0.1, 0.01, "0.00");

            if (MetricKey == "TVOC" && IsFixedFilterWideTvocMaterial(material))
                return new YAxisRule(0, 5.0, 0.5, "0.0");

            if (MetricKey == "PressureDrop" && IsFixedFilterPressureDropMaterial(material))
                return new YAxisRule(0, 35, 5, "0");

            if (MetricKey == "PressureDrop" && IsFixedFilterWidePressureDropMaterial(material))
                return new YAxisRule(0, 150, 10, "0");

            return null;
        }

        private bool IsFixedCylinderMaterial(string material)
        {
            return IsMaterial(material, "IKP201", "IKP205", "CI001", "SG017_A", "SG017_C", "SG017_D", "SG029", "SG035");
        }

        private bool IsFixedFilterParticleMaterial(string material)
        {
            return IsMaterial(
                material,
                "12G0A840000013",
                "12K0A840000015",
                "12K0A840000016",
                "12K0A840000017",
                "12B0P550000002",
                "12B0P550000009",
                "12B0P550000004",
                "12A0H660000001");
        }

        private bool IsFixedFilterTvocMaterial(string material)
        {
            return IsMaterial(material, "12G0A840000013", "12K0A840000015", "12K0A840000016", "12K0A840000017", "12B0P550000002");
        }

        private bool IsFixedFilterWideTvocMaterial(string material)
        {
            return IsMaterial(material, "12B0P550000009", "12B0P550000004", "12A0H660000001");
        }

        private bool IsFixedFilterPressureDropMaterial(string material)
        {
            return IsMaterial(material, "12K0A840000015", "12K0A840000016", "12K0A840000017");
        }

        private bool IsFixedFilterWidePressureDropMaterial(string material)
        {
            return IsMaterial(
                material,
                "12G0A840000013",
                "12B0P550000002",
                "12B0P550000009",
                "12B0P550000004",
                "12A0H660000001");
        }

        private bool IsCylinderFinishedChart(string material)
        {
            if (points.Count > 0)
                return HasSource("P5");

            return IsMaterial(material, "IKP201", "IKP205", "CI001", "SG017_A", "SG017_C", "SG017_D", "SG029", "SG035");
        }

        private bool IsFilterFinishedChart(string material)
        {
            if (points.Count > 0)
                return HasSource("P3");

            return material.StartsWith("12", StringComparison.OrdinalIgnoreCase);
        }

        private bool HasSource(string source)
        {
            return points.Any(x => string.Equals(x.Source, source, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsMaterial(string material, params string[] values)
        {
            return values.Any(x => string.Equals(material, NormalizeMaterial(x), StringComparison.OrdinalIgnoreCase));
        }

        private string NormalizeMaterial(string value)
        {
            return (value ?? "").Trim();
        }

        private void DrawEmptyMessage(Graphics g, Font font, Rectangle plot)
        {
            if (string.IsNullOrWhiteSpace(EmptyMessage))
                return;

            SizeF size = g.MeasureString(EmptyMessage, font);
            g.DrawString(
                EmptyMessage,
                font,
                Brushes.Gray,
                plot.Left + (plot.Width - size.Width) / 2F,
                plot.Top + (plot.Height - size.Height) / 2F);
        }

        private void DrawSamplePoints(Graphics g, Rectangle plot, double minY, double maxY)
        {
            using (var brush = new SolidBrush(Color.FromArgb(20, 158, 208)))
            {
                for (int i = 0; i < points.Count; i++)
                {
                    float x = MapX(plot, i);
                    float y = MapY(plot, minY, maxY, points[i].Value);
                    g.FillEllipse(brush, x - 4F, y - 4F, 8F, 8F);
                }
            }
        }

        private void DrawDateLabels(Graphics g, Font font, Rectangle plot)
        {
            if (points.Count == 0)
                return;

            const float minLabelSpacing = 72F;
            int maxLabelCount = Math.Max(2, Math.Min(12, (int)(plot.Width / minLabelSpacing)));
            int step = Math.Max(1, (int)Math.Ceiling(points.Count / (double)maxLabelCount));
            int lastIndex = points.Count - 1;
            var labelIndexes = new List<int>();

            for (int i = 0; i < points.Count; i += step)
            {
                if (i == lastIndex)
                    continue;

                if (labelIndexes.Count > 0 &&
                    points[labelIndexes[labelIndexes.Count - 1]].TestDate.Date == points[i].TestDate.Date)
                {
                    continue;
                }

                if (labelIndexes.Count > 0 &&
                    MapX(plot, i) - MapX(plot, labelIndexes[labelIndexes.Count - 1]) < minLabelSpacing)
                {
                    continue;
                }

                labelIndexes.Add(i);
            }

            while (labelIndexes.Count > 0)
            {
                int previousIndex = labelIndexes[labelIndexes.Count - 1];
                bool sameDate = points[previousIndex].TestDate.Date == points[lastIndex].TestDate.Date;
                bool tooClose = MapX(plot, lastIndex) - MapX(plot, previousIndex) < minLabelSpacing;

                if (!sameDate && !tooClose)
                    break;

                labelIndexes.RemoveAt(labelIndexes.Count - 1);
            }

            labelIndexes.Add(lastIndex);

            foreach (int index in labelIndexes)
                DrawDateLabel(g, font, plot, index);
        }

        private void DrawDateLabel(Graphics g, Font font, Rectangle plot, int index)
        {
            float x = MapX(plot, index);
            g.TranslateTransform(x - 8, plot.Bottom + 10);
            g.RotateTransform(90);
            g.DrawString(points[index].TestDate.ToString("yyyy.MM.dd"), font, Brushes.Black, 0, 0);
            g.ResetTransform();
        }

        private void DrawHorizontalLine(Graphics g, Rectangle plot, double minY, double maxY, double value, Pen pen)
        {
            float y = MapY(plot, minY, maxY, value);
            g.DrawLine(pen, plot.Left, y, plot.Right, y);
        }

        private double GetSigmaLineValue(int multiplier)
        {
            if (!average.HasValue || !sigma.HasValue)
                return 0;

            double value = average.Value + sigma.Value * multiplier;

            if ((MetricKey == "Particle" || MetricKey == "TVOC") && value < 0)
                return 0;

            return value;
        }

        private void DrawLegend(Graphics g, Font font, Rectangle plot)
        {
            int y = Math.Max(plot.Bottom + 116, Height - 82);
            int x = plot.Left - 10;
            int maxX = Math.Max(x + 260, plot.Right);

            using (var blueBrush = new SolidBrush(Color.FromArgb(20, 158, 208)))
            using (var redDash = new Pen(Color.Red, 2.5F) { DashStyle = DashStyle.Dash })
            using (var green = new Pen(Color.FromArgb(64, 160, 46), 2.5F))
            using (var red = new Pen(Color.Red, 2.5F))
            {
                x = DrawDotLegendItem(g, font, blueBrush, x, ref y, maxX, PointLegendText);

                if (average.HasValue && sigma.HasValue)
                {
                    if (ShowPlus1Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "+1σ=" + FormatNumber(GetSigmaLineValue(1)));

                    if (ShowMinus1Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "-1σ=" + FormatNumber(GetSigmaLineValue(-1)));

                    if (ShowPlus2Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "+2σ=" + FormatNumber(GetSigmaLineValue(2)));

                    if (ShowMinus2Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "-2σ=" + FormatNumber(GetSigmaLineValue(-2)));

                    if (ShowPlus3Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "+3σ=" + FormatNumber(GetSigmaLineValue(3)));

                    if (ShowMinus3Sigma)
                        x = DrawLineLegendItem(g, font, redDash, x, ref y, maxX, "-3σ=" + FormatNumber(GetSigmaLineValue(-3)));
                }

                if (ShowAverage && average.HasValue)
                    x = DrawLineLegendItem(g, font, green, x, ref y, maxX, "平均值=" + FormatNumber(average.Value));

                if (ShowLsl && lsl.HasValue)
                    x = DrawLineLegendItem(g, font, red, x, ref y, maxX, "LSL=" + FormatNumber(lsl.Value));

                if (ShowUsl && usl.HasValue)
                    DrawLineLegendItem(g, font, red, x, ref y, maxX, "USL=" + FormatNumber(usl.Value));
            }
        }

        private int DrawDotLegendItem(Graphics g, Font font, Brush brush, int x, ref int y, int maxX, string text)
        {
            int width = 32 + (int)g.MeasureString(text, font).Width + 22;

            if (x + width > maxX)
            {
                x = 86;
                y += 28;
            }

            g.FillEllipse(brush, x, y + 6, 8, 8);
            g.DrawString(text, font, Brushes.Black, x + 22, y);

            return x + width;
        }

        private int DrawLineLegendItem(Graphics g, Font font, Pen pen, int x, ref int y, int maxX, string text)
        {
            int width = 50 + (int)g.MeasureString(text, font).Width + 22;

            if (x + width > maxX)
            {
                x = 86;
                y += 28;
            }

            g.DrawLine(pen, x, y + 10, x + 34, y + 10);
            g.DrawString(text, font, Brushes.Black, x + 44, y);

            return x + width;
        }

        private float MapX(Rectangle plot, int index)
        {
            if (points.Count <= 1)
                return plot.Left + plot.Width / 2F;

            return plot.Left + plot.Width * index / (float)(points.Count - 1);
        }

        private float MapY(Rectangle plot, double minY, double maxY, double value)
        {
            return (float)(plot.Bottom - (value - minY) / (maxY - minY) * plot.Height);
        }

        private string FormatNumber(double value)
        {
            double abs = Math.Abs(value);

            if (abs >= 1000)
                return value.ToString("0");

            if (abs >= 10)
                return value.ToString("0.0");

            return value.ToString("0.00");
        }
    }
}

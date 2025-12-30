using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

public enum CalcMode { Moisture, Ash , NButane}

public class FormCalcMoistureAsh : Form
{
    private readonly CalcMode _mode;

    // 3 組並排
    TextBox[] tbCrucible = new TextBox[3];
    TextBox[] tbSample = new TextBox[3];
    TextBox[] tbTotal = new TextBox[3];
    TextBox[] tbResult = new TextBox[3];

    TextBox tbAverage;
    Label lblMode;
    Button btnOK, btnCancel;

    public double? AverageResult { get; private set; }
    public double[] ColumnResults { get; private set; } = new double[3];

    public FormCalcMoistureAsh(CalcMode mode)
    {
        _mode = mode;
        InitUI();
    }

    private void InitUI()
    {
        Text = _mode == CalcMode.Moisture ? "含水率計算" : "灰分計算";
        Font = new Font("Microsoft JhengHei UI", 10F);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(380,380);
        MaximizeBox = MinimizeBox = false;

        lblMode = new Label
        {
            AutoSize = false,
            TextAlign = ContentAlignment.MiddleLeft,
            Text = _mode == CalcMode.Moisture ? "模式：含水率(%)" : "模式：灰分(%)",
            Bounds = new Rectangle(10, 10, 320, 22)
        };
        Controls.Add(lblMode);

        // 欄寬與定位
        int x0 = 10, wLabel = 70, wBox = 80, h = 26, gapX = 8, gapY = 8;
        int c1 = x0 + wLabel + gapX;         // 第1欄 textbox x
        int c2 = c1 + wBox + gapX;           // 第2欄
        int c3 = c2 + wBox + gapX;           // 第3欄
        int y = 40;

        // 工具：建立一行
        void addRow(string label, TextBox[] tbs, bool readOnly = false)
        {
            var lab = new Label { Text = label, Bounds = new Rectangle(x0, y, wLabel, h), TextAlign = ContentAlignment.MiddleLeft };
            Controls.Add(lab);

            for (int i = 0; i < 3; i++)
            {
                var tb = new TextBox
                {
                    Bounds = new Rectangle((i == 0 ? c1 : i == 1 ? c2 : c3), y, wBox, h),
                    TextAlign = HorizontalAlignment.Right,
                    ReadOnly = readOnly
                };
                if (!readOnly) tb.TextChanged += (_, __) => RecalcAll();
                tbs[i] = tb;
                Controls.Add(tb);
            }
            y += h + gapY;
        }

        // 依序加入行
        var dummy = new TextBox[3]; // 坩鍋編號若要輸入可改成 TextBox，現在省略
        addRow("坩鍋編號", dummy);                // 你要真的填編號就把 dummy 改成 tbCrucibleId[]
        addRow("空重", tbCrucible);
        addRow("樣品重", tbSample);
        addRow("總重", tbTotal);
        addRow("結果(%)", tbResult, readOnly: true);

        // 平均
        var labAvg = new Label { Text = "平均(%)", Bounds = new Rectangle(x0, y, wLabel, h), TextAlign = ContentAlignment.MiddleLeft };
        Controls.Add(labAvg);
        tbAverage = new TextBox { Bounds = new Rectangle(c1, y, (wBox * 3 + gapX * 2), h), ReadOnly = true, TextAlign = HorizontalAlignment.Right };
        Controls.Add(tbAverage);
        y += h + gapY;

        // 說明（公式）
        var lblHint = new Label
        {
            AutoSize = false,
            TextAlign = ContentAlignment.TopLeft,
            Bounds = new Rectangle(x0, y, c3 + wBox - x0, 36),
            ForeColor = Color.DimGray
        };
        if (_mode == CalcMode.Ash)
            lblHint.Text = "灰分(%) = (總重 - 空重) / 樣品重 × 100";
        else
            lblHint.Text = "含水率(%) ≈ (總重 - 空重) / 樣品重 × 100";
        Controls.Add(lblHint);
        y += 40;

        // 按鈕
        btnOK = new Button { Text = "確定", Bounds = new Rectangle(c2 - 40, y, 80, 28), DialogResult = DialogResult.OK };
        btnCancel = new Button { Text = "取消", Bounds = new Rectangle(c3 - 40, y, 80, 28), DialogResult = DialogResult.Cancel };
        btnOK.Click += (_, __) => { RecalcAll(); AverageResult = ParseD(tbAverage.Text); Close(); };
        btnCancel.Click += (_, __) => { AverageResult = null; Close(); };
        Controls.AddRange(new Control[] { btnOK, btnCancel });

        AcceptButton = btnOK;
        CancelButton = btnCancel;
    }

    // 解析工具
    private static double? ParseD(string s) => double.TryParse(s, out var v) ? v : (double?)null;

    // 計算單一欄
    private double? CalcOne(int i)
    {
        var c = ParseD(tbCrucible[i].Text);
        var s = ParseD(tbSample[i].Text);
        var t = ParseD(tbTotal[i].Text);
        if (c == null || s == null || t == null || s == 0) return null;

        double result;
        if (_mode == CalcMode.Ash)
        {
            // 灰分(%) = (總重 - 空重) / 樣品重 × 100
            result = ((t.Value - c.Value) / s.Value) * 100.0;
        }
        else
        {
            // 含水率(%) 近似： (樣品重 - 乾重) / 樣品重 × 100
            // 乾重 ≈ (總重 - 空重)
            double dry = (t.Value - c.Value);
            result = ((s.Value - dry) / s.Value) * 100.0;
        }
        return result;
    }

    private void RecalcAll()
    {
        double sum = 0; int cnt = 0;
        for (int i = 0; i < 3; i++)
        {
            var v = CalcOne(i);
            ColumnResults[i] = double.NaN;
            if (v == null)
            {
                tbResult[i].Text = "";
            }
            else
            {
                double val = v.Value;
                tbResult[i].Text = val.ToString("F2");
                ColumnResults[i] = val;
                sum += val; cnt++;
            }
        }
        tbAverage.Text = (cnt > 0) ? (sum / cnt).ToString("F2") : "";
    }
}

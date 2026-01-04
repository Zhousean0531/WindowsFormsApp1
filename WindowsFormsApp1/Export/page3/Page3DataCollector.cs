using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public static class Page3DataCollector
{
    public static Page3ExportData Collect(TabPage tab)
    {
        var testPicker = ControlHelper.Find<DateTimePicker>(tab, "SemiTestDateBox");
        var arrivePicker = ControlHelper.Find<DateTimePicker>(tab, "SemiArriveDateBox");
        var materialBox = ControlHelper.Find<ComboBox>(tab, "SemiTypeBox");

        if (testPicker == null || arrivePicker == null || materialBox == null)
        {
            MessageBox.Show("缺少必要欄位");
            return null;
        }

        string testDate = testPicker.Value.ToString("yyyy.MM.dd");
        string arriveDate = arrivePicker.Value.ToString("yyyy.MM.dd");
        string material = materialBox.Text.Trim();

        var lotNos = StringUtil.Split(
            ControlHelper.GetText(tab, "SemiLotBox")
        );

        var weights = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "SemiWeightBox")
        );

        var vocIn = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "SemiVocInBox")
        );

        var vocOut = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "SemiVocOutBox")
        );

        var deltas = StringUtil.SplitDouble(
            ControlHelper.GetText(tab, "SemiPressureBox")
        );

        int n = new[] {
            lotNos.Count,
            weights.Count,
            vocIn.Count,
            vocOut.Count,
            deltas.Count
        }.Min();

        if (n <= 0)
        {
            MessageBox.Show("沒有可匯出的資料");
            return null;
        }

        lotNos = lotNos.Take(n).ToList();
        weights = weights.Take(n).ToList();
        vocIn = vocIn.Take(n).ToList();
        vocOut = vocOut.Take(n).ToList();
        deltas = deltas.Take(n).ToList();

        var outMinusIn = vocOut.Zip(vocIn, (o, i) =>
        {
            var d = o - i;
            return d <= 0 ? "N.D." : d.ToString("F1");
        }).ToList();

        return new Page3ExportData
        {
            TestingDate = testDate,
            ArrivalDate = arriveDate,
            Material = material,

            LotNos = lotNos,
            Weights = weights,
            DeltaPs = deltas,
            VocIns = vocIn,
            VocOuts = vocOut,
            OutMinusIn = outMinusIn
        };
    }
}

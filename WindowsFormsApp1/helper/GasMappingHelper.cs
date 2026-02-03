using System;
using System.Collections.Generic;
using static WindowsFormsApp1.Helpers.GasMappingHelper;

namespace WindowsFormsApp1.Helpers
{
    public static class GasMappingHelper
    {
        public enum GasPageType
        {
            FilterInProcess,   // TabPage2
            CylinderRawPage    // TabPage4
        }

        /// 依原料型號判斷對應的測試氣體
        public static string GetGasNameFromRawType(string rawType)
        {
            if (string.IsNullOrWhiteSpace(rawType))
                return string.Empty;

            rawType = rawType.Trim().ToUpperInvariant();

            // SG 系列
            if (rawType.Contains("SG"))
                return "Toluene";

            // IKP 系列
            if (rawType.StartsWith("IKP"))
                return "H2S";

            // 特定型號
            if (rawType == "SI013")
                return "SO2";

            // 找不到對應
            return string.Empty;
        }
        private static readonly Dictionary<string, GasConfig> _map = new Dictionary<string, GasConfig>(StringComparer.OrdinalIgnoreCase)
        {
            ["SO2"] = new GasConfig
            {
                GasKey = "SO2",
                DisplayName = "SO₂",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInSO2CheckBox",
                        ConcBox = "FilterInSO2ConcentrationBox",
                        BgBox = "FilterInBackGroundSO2Box",
                        ValueBox = "FilterInValueSO2Box"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylinderSO2CheckBox",
                        ConcBox = "CylinderRawSO2ConcertrationBox",
                        BgBox = "CylinderRawSO2BackGroundBox",
                        ValueBox = "CylinderRawSO2ValueBox"
                    }
                }
            },
            ["H2S"] = new GasConfig
            {
                GasKey = "H2S",
                DisplayName = "H₂S",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInH2SCheckBox",
                        ConcBox = "FilterInH2SConcentrationBox",
                        BgBox = "FilterInBackGroundH2SBox",
                        ValueBox = "FilterInValueH2SBox"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylinderH2SCheckBox",
                        ConcBox = "CylinderRawH2SConcertrationBox",
                        BgBox = "CylinderRawH2SBackGroundBox",
                        ValueBox = "CylinderRawH2SValueBox"
                    }
                }
            },
            ["Toluene"] = new GasConfig
            {
                GasKey = "Toluene",
                DisplayName = "Toluene",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInTolueneCheckBox",
                        ConcBox = "FilterInTolueneConcentrationBox",
                        BgBox = "FilterInBackGroundTolueneBox",
                        ValueBox = "FilterInValueTolueneBox"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylinderRawTolueneCheckBox",
                        ConcBox = "CylinderRawTolueneConcertrationBox",
                        BgBox = "CylinderRawTolueneBackGroundBox",
                        ValueBox = "CylinderRawTolueneValueBox"
                    }
                }
            },
            ["NH3"] = new GasConfig
            {
                GasKey = "NH3",
                DisplayName = "NH₃",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInNH3CheckBox",
                        ConcBox = "FilterInNH3ConcentrationBox",
                        BgBox = "FilterInBackGroundNH3Box",
                        ValueBox = "FilterInValueNH3Box"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylinderNH3CheckBox",
                        ConcBox = "CylinderRawNH3ConcertrationBox",
                        BgBox = "CylinderRawNH3BackGroundBox",
                        ValueBox = "CylinderRawNH3ValueBox"
                    }
                }
            },
            ["IPA"] = new GasConfig
            {
                GasKey = "IPA",
                DisplayName = "IPA",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInIPACheckBox",
                        ConcBox = "FilterInIPAConcentrationBox",
                        BgBox = "FilterInBackGroundIPABox",
                        ValueBox = "FilterInValueIPABox"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylinderRawIPACheckBox",
                        ConcBox = "CylinderRawIPAConcertrationBox",
                        BgBox = "CylinderRawIPABackGroundBox",
                        ValueBox = "CylinderRawIPAValueBox"
                    }
                }
            },
            ["Acetone"] = new GasConfig
            {
                GasKey = "Acetone",
                DisplayName = "Acetone",
                UiMap = new Dictionary<GasPageType, GasUiConfig>
                {
                    [GasPageType.FilterInProcess] = new GasUiConfig
                    {
                        CheckBoxName = "FilterInAcetoneCheckBox",
                        ConcBox = "FilterInAcetoneConcentrationBox",
                        BgBox = "FilterInBackGroundAcetoneBox",
                        ValueBox = "FilterInValueAcetoneBox"
                    },
                    [GasPageType.CylinderRawPage] = new GasUiConfig
                    {
                        CheckBoxName = "CylindeRawAcetoneCheckBox",
                        ConcBox = "CylinderRawAcetoneConcertrationBox",
                        BgBox = "CylinderRawAcetoneBackGroundBox",
                        ValueBox = "CylinderRawAcetoneValueBox"
                    }
                }
            },
        };
        public static GasConfig Get(string gasKey)
        => gasKey != null && _map.TryGetValue(gasKey, out var cfg) ? cfg : null;
    }
}
public class GasUiConfig
{
    public string CheckBoxName { get; set; }
    public string ConcBox { get; set; }
    public string BgBox { get; set; }
    public string ValueBox { get; set; }
}

public class GasConfig
{
    public string GasKey { get; set; }
    public string DisplayName { get; set; }

    // ⭐ 關鍵：每個頁面一組 UI 對應
    public Dictionary<GasPageType, GasUiConfig> UiMap { get; set; }
}

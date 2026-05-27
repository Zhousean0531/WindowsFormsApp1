using System;
using System.Collections.Generic;

namespace WindowsFormsApp1.Data_Access.Analysis
{
    public class QualityAnalysisPoint
    {
        public DateTime TestDate { get; set; }
        public double Value { get; set; }
        public string Source { get; set; }
        public int? BatchId { get; set; }
        public int? RowNo { get; set; }
        public string GroupKey { get; set; }
    }

    public class QualityAnalysisSetting
    {
        public string Material { get; set; }
        public string MetricName { get; set; }
        public string SigmaMode { get; set; }
        public int? SigmaMonths { get; set; }
        public DateTime? SigmaStartDate { get; set; }
        public DateTime? SigmaEndDate { get; set; }
        public double? USL { get; set; }
        public double? LSL { get; set; }
    }

    public class QualityAnalysisResult
    {
        public List<QualityAnalysisPoint> Points { get; set; } = new List<QualityAnalysisPoint>();
        public List<QualityAnalysisPoint> SigmaPoints { get; set; } = new List<QualityAnalysisPoint>();
        public QualityAnalysisSetting Setting { get; set; }
        public DateTime SigmaStartDate { get; set; }
        public DateTime SigmaEndDate { get; set; }
        public double? Average { get; set; }
        public double? Sigma { get; set; }
    }
}

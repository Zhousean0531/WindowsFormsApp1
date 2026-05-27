using System;

namespace WindowsFormsApp1.Data_Access.Query
{
    public class QcQueryCriteria
    {
        public string QueryKind { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string RawMaterialType { get; set; }
        public string SemiProductType { get; set; }
        public string SemiProductGsm { get; set; }
        public string SemiProductMaterialNo { get; set; }
        public string ProductNo { get; set; }
    }
}

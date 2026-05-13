using System;
using System.Collections.Generic;

namespace WindowsFormsApp1.Data_Access.Page1
{
    public class P1Batch
    {
        public string ReportNo { get; set; }
        public string Material { get; set; }
        public decimal? Concentration { get; set; }
        public decimal? Background { get; set; }
        public string MaterialNo { get; set; }
        public string QtyText { get; set; }
        public DateTime ArrivalDate { get; set; }
        public DateTime TestingDate { get; set; }
        public string Username { get; set; }
        public Dictionary<string, double> ParticleSizePercentages { get; set; }
            = new Dictionary<string, double>();

        public List<P1Sample> Samples { get; set; }
            = new List<P1Sample>();
    }

    public class P1Sample
    {
        public string LotFull { get; set; }
        public string SuppliedNO { get; set; }

        public decimal? Weight { get; set; }
        public decimal? Density { get; set; }
        public decimal? DeltaP { get; set; }
        public decimal? VocIn { get; set; }
        public decimal? VocOut { get; set; }
        public decimal? Outgassing { get; set; }

        public List<decimal?> Efficiencies { get; set; }

        public bool IsSelected { get; set; }

    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Data_Access.Page1
{
    public class P1Batch
    {
        public string ReportNo { get; set; }
        public string Material { get; set; }
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
        public double? Weight { get; set; }
        public double? Density { get; set; }
        public double? DeltaP { get; set; }
        public double? VocIn { get; set; }
        public double? VocOut { get; set; }
        public string Outgassing { get; set; }

        public bool IsSelected { get; set; }

        public List<double> Efficiencies { get; set; }
            = new List<double>();
    }
}
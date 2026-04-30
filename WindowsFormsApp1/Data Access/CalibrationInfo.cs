using System;
namespace WindowsFormsApp1.Data_Access
{
    public class CalibrationInfo
    {
        public int Id { get; set; }
        public string InstrumentName { get; set; }
        public DateTime CalibrationDate { get; set; }
        public DateTime ExpireDate { get; set; }
    }
}
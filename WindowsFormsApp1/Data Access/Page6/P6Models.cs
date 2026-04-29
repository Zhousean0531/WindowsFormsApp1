using System;
using System.Collections.Generic;

namespace WindowsFormsApp1.Data_Access.Page6
{
    public class P6Batch
    {
        public string ReportNo { get; set; }
        public DateTime TestDate { get; set; }
        public string UserName { get; set; }

        public List<P6Item> Items { get; set; } = new List<P6Item>();
    }
}
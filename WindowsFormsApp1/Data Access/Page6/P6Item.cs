namespace WindowsFormsApp1.Data_Access.Page6
{
    public class P6Item
    {
        public string Col1 { get; set; }  // DGV[0]
        public string Col2 { get; set; }  // DGV[1]

        public string Spec1 { get; set; } // DGV[2]
        public string Spec2 { get; set; } // DGV[3]

        public string Range1 { get; set; } // DGV[4]
        public string Range2 { get; set; } // DGV[5]

        public string Result { get; set; } // DGV[6]
        public string Judgment { get; set; } // DGV[7]

        // 額外欄位（8之後）
        public string Extra1 { get; set; }
        public string Extra2 { get; set; }
        public string Extra3 { get; set; }

        public string Supplier { get; set; }
    }
}

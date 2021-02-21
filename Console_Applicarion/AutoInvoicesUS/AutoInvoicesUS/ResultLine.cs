using System;

namespace AutoInvoicesUS
{
    class ResultLine
    {
        public int TTP_Numerator { get; set; }
        public string Customer_Name { get; set; }
        public int Priority_Lines_Count { get; set; }
        public short Succeeded { get; set; }
        public int? Priority_Numerator { get; set; }
        public string Message { get; set; }
    }
}

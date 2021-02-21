using System;

namespace AutoInvoicesUS
{
    class FileStorageLine
    {
        public int TTP_Numerator { get; set; }
        public string File_Name_Desc { get; set; }
        public byte[] File_Stream { get; set; }
    }
}

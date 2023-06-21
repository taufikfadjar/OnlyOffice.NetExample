using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OnlyOfficePenagihanHutang.Models
{
    public class FileModel
    {
        public string BaseFileName { get; set; }
        public string Version { get; set; }
        public byte[] Bytes { get; set; }
        public string Mime { get; set; }
        public DateTime? LastModified { get; set; }

    }

    public class FileInfoOnlyOffice
    {
        public string BaseFileName { get; set; }
        public string Version { get; set; }
        public bool UserCanWrite { get; set; }
        public bool ReadOnly { get; set; }
        public bool SupportsUpdate { get; set; }
    }
}
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OnlyOfficePenagihanHutang.Helper
{
    public class WordReplacement
    {
        public string TextToReplace { get; set; }
        public string ReplacementText { get; set; }
        public bool IsCheckBox { get; set; }
        public bool MatchWholeText { get; set; }
        public List<Run> Checkboxes { get; set; }
        public Run Run { get; set; }
        public bool UseRun { get; set; }

        public WordReplacement(string toReplace, string replacement)
        {
            this.MatchWholeText = false;
            this.TextToReplace = toReplace;
            this.ReplacementText = replacement;
            this.UseRun = false;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GuidelinesExtractor
{
    public class Guideline
    {
        public string Key { get; set; }
        public string Text { get; set; }

        public string Severity { get; set; }

        public string Section { get; set; }
        public string Subsection { get; set; }
    }
}

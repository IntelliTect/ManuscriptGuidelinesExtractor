using System.Collections.Generic;


namespace GuidelinesExtractor
{
    public class Guideline
    {
        public string Key { get; set; } = "";
        public string Text { get; set; } = "";

        public string Severity { get; set; }= "";

        public string Section { get; set; }= "";
        public string Subsection { get; set; }= "";

        public List<string> Comments { get; set; } = new List<string>();

    }

}

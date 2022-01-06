using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GuidelinesExtractorTests
{
    [TestClass]
    public class FormatTests
    {
        [TestMethod][Ignore]
        public void TestMethod1()
        {
            string folder = @"C:\Users\PATH-TO\GuidelinesExtractor\GuidelineExtractorTests\BookmarkFiles\EssentialCSharpManuscript";
            string date = System.DateTime.Now.ToString("dd-MM-yy-H-mm-ss");


            // GuidelinesExtractor.GuidelinesFormatter.AllGuidelinesToCSVWithBookmarks(folder, verbose: true, guidelineTitleStyle: "SF2_TTL");

        }



        [TestMethod][Ignore]
        public void XMLOutputFileFormatAsExpected()
        {


            string s =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
	<guideline key=""Ch01_fa67753"" severity="""" section="""" subsection="""">DO favor clarity over brevity when naming identifiers.</guideline>
	<guideline key=""Ch01_674fe1e"" severity="""" section="""" subsection="""">DO NOT use abbreviations or contractions within identifier names.</guideline>
	<guideline key=""Ch01_024aa0c"" severity="""" section="""" subsection="""">DO NOT use any acronyms unless they are widely accepted, and even then use them consistently.</guideline>
</root>";

            string doc = @"PATH-TO/TestDoc.docx";

            //GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc, guidelineTitleStyle: "SF2_TTL");

        }
    }
}

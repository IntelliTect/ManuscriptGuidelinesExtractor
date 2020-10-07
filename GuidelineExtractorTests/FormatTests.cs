using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace GuidelineExtractorTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string folder = @"C:\Users\saffron\source\repos\GuidelinesExtractor\GuidelineExtractorTests\BookmarkFiles\EssentialCSharpManuscript";

            GuidelinesExtractor.GuidelinesFormatter.AllGuidelinesToCSVWithBookmarks(folder, verbose: true, guidelineTitleStyle: "SF2_TTL");

        }



        [TestMethod]
        public void XMLOutputFileFormatAsExpected()
        {


string s = 
@"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
	<guideline key=""Ch01_fa67753"" severity="""" section="""" subsection="""">DO favor clarity over brevity when naming identifiers.</guideline>
	<guideline key=""Ch01_674fe1e"" severity="""" section="""" subsection="""">DO NOT use abbreviations or contractions within identifier names.</guideline>
	<guideline key=""Ch01_024aa0c"" severity="""" section="""" subsection="""">DO NOT use any acronyms unless they are widely accepted, and even then use them consistently.</guideline>
</root>";

            string doc = @"C:\Users\saffron\source\repos\GuidelinesExtractor\Michaelis_Ch12.docx";

           // GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc , guidelineTitleStyle: "SF2_TTL");

        }
    }
}

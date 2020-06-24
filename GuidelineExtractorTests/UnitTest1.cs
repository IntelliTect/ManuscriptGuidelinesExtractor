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
        public void GetUniqueGuidelinesInTable()
        {
            string doc = @"C:\Users\saffron\source\repos\GuidelinesExtractor\Michaelis_Ch12.docx";

           // GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc , guidelineTitleStyle: "SF2_TTL");

        }
    }
}

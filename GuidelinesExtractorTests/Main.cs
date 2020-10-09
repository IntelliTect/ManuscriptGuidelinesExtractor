using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GuidelinesExtractor;


namespace GuidelinesExtractorTests
{
    [TestClass]
    public class Main
    {
        [TestMethod]
        public void TestMethod1()
        {
            string folder = @"C:\Users\saffron\source\repos\EssentialCSharpManuscript\GuidelinesExtractor\WordDocs";

            GuidelinesFormatter.AllGuidelinesToXML(folder,false,"SF2_TTL", WordDocGuidelineTools.ExtractionMode.BookmarkOnlyNewGuidelinesAndCheckForChangesOfPreviouslyBookmarkedGuidelines);
            //GuidelinesExtractor (folder, verbose: true, guidelineTitleStyle: "SF2_TTL");



        }



        [TestMethod]
        public void GetUniqueGuidelinesInTable()
        {
            string doc = @"C:\Users\saffron\source\repos\GuidelinesExtractor\Michaelis_Ch12.docx";

            // GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc , guidelineTitleStyle: "SF2_TTL");

        }
    }
}

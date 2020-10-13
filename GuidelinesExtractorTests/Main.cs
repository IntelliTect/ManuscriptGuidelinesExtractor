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
            string folder = @"C:\Users\saffron\Downloads\FinalCopyCSharp8";
            string pathToExistingGuidelines = @"C:\Users\saffron\source\repos\EssentialCSharpManuscript\GuidelinesExtractor\WordDocs\Guidelines10 - 10 - 20.xml";


            GuidelinesFormatter guidelinesFormatter = new GuidelinesFormatter(folder, "SF2_TTL", WordDocGuidelineTools.ExtractionMode.BookmarkAllGuidelines);

            //GuidelinesFormatter guidelinesFormatter = new GuidelinesFormatter(folder, "SF2_TTL", WordDocGuidelineTools.ExtractionMode.BookmarkOnlyNewGuidelinesAndCheckForChangesOfPreviouslyBookmarkedGuidelines, pathToExistingGuidelines);
            guidelinesFormatter.AllGuidelinesToXML();
            



        }



        [TestMethod]
        public void GetUniqueGuidelinesInTable()
        {
            //string doc = @"C:\Users\saffron\source\repos\GuidelinesExtractor\Michaelis_Ch12.docx";

            // GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc , guidelineTitleStyle: "SF2_TTL");

        }
    }
}

using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GuidelinesExtractor;


namespace GuidelinesExtractorTests
{
    [TestClass]
    public class RUNPROGRAM
    {
        [TestMethod]
        public void Run_Program()
        {
            // REPLACE THESE with the full path to the folder containing the word doc
            string wordDocFolder = @"C:/Users/PATH-TO/ManuscriptGuidelinesExtractor/WordDocs";
            string pathToExistingGuidelines = @"C:/Users/PATH-TO/ManuscriptGuidelinesExtractor/Guidelines.xml";

            // GuidelinesFormatter(string pathToChapterDocumentFolder, string guidelineTitleStyle, WordDocGuidelineTools.ExtractionMode extractionMode, string pathToExistingGuidelinesXml = null)
            // The Extraction mode can be changed to compare to existing .xml file
            GuidelinesFormatter guidelinesFormatter = new GuidelinesFormatter(wordDocFolder, "SF2_TTL", WordDocGuidelineTools.ExtractionMode.BookmarkAllGuidelines);
            guidelinesFormatter.AllGuidelinesToXML();
        }
    }
}

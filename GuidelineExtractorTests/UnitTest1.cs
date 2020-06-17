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
            string folder = @"C:\Users\saffron\source\repos\GuidelinesExtractor\GuidelineExtractorTests\Quick Test Files";

            GuidelinesExtractor.GuidelinesFormatter.AllGuidelinesToMarkDown(folder, verbose: true);

        }



        [TestMethod]
        public void GetUniqueGuidelinesInTable()
        {
            string doc = @"C:\Users\saffron\source\repos\GuidelinesExtractor\GuidelineExtractorTests\Quick Test Files";

            GuidelinesExtractor.GuideLineTools.GetGuideLinesInDocument(doc);

        }
    }
}

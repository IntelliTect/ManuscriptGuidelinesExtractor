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
            string folder = @"C:\Users\saffron\source\repos\GuidelinesExtractor\GuidelinesExtractor\Quick Test Files";

            GuidelinesExtractor.GuidelinesFormatter.AllGuidelinesToMarkDown(folder, verbose: true);

        }
    }
}

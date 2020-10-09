using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace GuidelinesExtractor
{
    public class GuidelinesFormatter
    {
        public const string _Guideline = "guideline";
        public const string _Key= "key";
         public const string _Severity  = "severity";

         public const string _Section  = "section";
         public const string _Subsection  = "subsection";
        public const string _Comments = "comments";



        /// <summary>
        /// Will extract all guidelines (that match style) from documents in folder and put them in a xml file in that folder.
        /// </summary>
        public static void AllGuidelinesToXML(string pathToChapterDocumentFolder, bool verbose, string guidelineTitleStyle, WordDocGuidelineTools.ExtractionMode extractionMode)
        {

            List<string> allDocs = FileManager.GetAllFilesAtPath(pathToChapterDocumentFolder, searchPattern: "Michaelis_Ch??.docx")
                .OrderBy(x => x).ToList();

            int chapterNumber = 0;
            List<Guideline> allChapterGuidelines = new List<Guideline>();
            foreach (string chapterDocxPath in allDocs)
            {
                chapterNumber = GetChapterNumber(chapterDocxPath);
                //instead of the following line using GetGuideLinesInDocument. a new method called GetUniqueGuideLinesInDocument(exiting guidelines stored in xml) 
                //can be used to just find new guidelines (by determining if the guideline has a bookmark on it(note bookmarks only cover 255 characters)) and append them to the xml file 
                allChapterGuidelines.AddRange(WordDocGuidelineTools.GetGuideLinesInDocument(chapterDocxPath, chapterNumber, guidelineTitleStyle , extractionMode));

            }
             WriteXML(allChapterGuidelines, pathToChapterDocumentFolder);
            
            WordDocGuidelineTools._WordApp.Quit();

        }

        private static void WriteXML(List<Guideline> currentChapterGuidelines, string pathToChapterDocumentFolder)
        {
            XDocument doc = new XDocument(new XElement("Guidelines"));
            
            
            // example format
            // <guideline key="Ch01_fa67753" severity="DO" section="Naming" subsection="Variables and fields">DO favor clarity over brevity when naming identifiers.</guideline>
            foreach (Guideline guideline in currentChapterGuidelines)
            {
                var newGuideline = new XElement(_Guideline);
                newGuideline.SetAttributeValue(_Key, guideline.Key);
                newGuideline.SetAttributeValue(_Severity, guideline.Severity);
                newGuideline.SetAttributeValue(_Section, guideline.Section);
                newGuideline.SetAttributeValue(_Subsection, guideline.Subsection);
                newGuideline.SetValue(guideline.Text);
                //if (guideline.Comments.Count > 0) newGuideline.Add(guideline.Comments); add the comments as a child
                doc.Root.Add(newGuideline);
            }
            string date = System.DateTime.Today.ToString("dd - MM - yy");
            doc.Save(pathToChapterDocumentFolder + @"\Guidelines" + date + ".xml");

        }

        public static int GetChapterNumber(string filePath)
        {
            Regex regex = new Regex(@"Michaelis_Ch(\d{2}).docx$");

            var matches = regex.Match(filePath);

            if (int.TryParse(matches.Groups[1].Value, out int chapterNumber)
                && matches.Success)
            {
                return chapterNumber;
            }
            else
            {
                throw new Exception($"Cannot parse chapter number from {filePath}");
            }
        }

    }
}

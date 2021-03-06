using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace GuidelinesExtractor
{
    public class GuidelinesFormatter
    {
        public const string _Guideline = "guideline";
        public const string _Key = "key";
        public const string _Severity = "severity";

        public const string _Section = "section";
        public const string _Subsection = "subsection";
        public const string _Comments = "comments";



        public static HashSet<Guideline> Guidelines { get; private set; } = new HashSet<Guideline>();

        public string PathToChapterDocumentFolder { get; }
        public string GuidelineTitleStyle { get; }
        public WordDocGuidelineTools.ExtractionMode ExtractionMode { get; }
        public string PathToExistingGuidelinesXml { get; }

        public GuidelinesFormatter(string pathToChapterDocumentFolder, string guidelineTitleStyle, WordDocGuidelineTools.ExtractionMode extractionMode, string pathToExistingGuidelinesXml = null)
        {
            PathToChapterDocumentFolder = pathToChapterDocumentFolder;
            GuidelineTitleStyle = guidelineTitleStyle;
            ExtractionMode = extractionMode;

            if (ExtractionMode != WordDocGuidelineTools.ExtractionMode.BookmarkAllGuidelines && string.IsNullOrEmpty(pathToExistingGuidelinesXml))
            {
                throw new ArgumentNullException("pathToExistingGuidelinesXml not specified");
            }
            else if (ExtractionMode != WordDocGuidelineTools.ExtractionMode.BookmarkAllGuidelines)
            {
                Guidelines = ReadExisitingGuidelinesFile(pathToExistingGuidelinesXml);
                PathToExistingGuidelinesXml = pathToExistingGuidelinesXml;
            }
        }

        private HashSet<Guideline> ReadExisitingGuidelinesFile(string pathToExistingGuidelinesXml)
        {
            
            XDocument previousGuidelines = XDocument.Load(pathToExistingGuidelinesXml);

            HashSet<Guideline> guidelines  = new HashSet<Guideline>();

            foreach (var guidelineFromXml in previousGuidelines.Root.DescendantNodes().OfType<XElement>()) {
                Guideline guideline = new Guideline();
                guideline.Severity= guidelineFromXml.Attribute(_Severity)?.Value;
                guideline.Subsection = guidelineFromXml.Attribute(_Subsection)?.Value;
                guideline.Section = guidelineFromXml.Attribute(_Section)?.Value;
                guideline.Text = guidelineFromXml?.Value;
                guideline.Key= guidelineFromXml.Attribute(_Key)?.Value;

                guidelines.Add(guideline);
            }
            return guidelines;
        }




        /// <summary>
        /// Will extract all guidelines (that match style) from documents in folder and put them in a xml file in that folder.
        /// </summary>
        public void AllGuidelinesToXML()
        {

            List<string> allDocs = FileManager.GetAllFilesAtPath(PathToChapterDocumentFolder, searchPattern: "Michaelis_Ch??.docx")
                .OrderBy(x => x).ToList();

            int chapterNumber = 0;
            List<Guideline> allChapterGuidelines = new List<Guideline>();
            foreach (string chapterDocxPath in allDocs)
            {
                chapterNumber = GetChapterNumber(chapterDocxPath);      
                allChapterGuidelines.AddRange(WordDocGuidelineTools.GetGuideLinesInDocument(chapterDocxPath, chapterNumber, GuidelineTitleStyle, ExtractionMode));

            }
            WriteXML(PathToChapterDocumentFolder);

            WordDocGuidelineTools._WordApp.Quit();

        }

        private void WriteXML( string pathToChapterDocumentFolder)
        {
            XDocument doc = new XDocument(new XElement("Guidelines"));


            // example format
            // <guideline key="Ch01_fa67753" severity="DO" section="Naming" subsection="Variables and fields">DO favor clarity over brevity when naming identifiers.</guideline>
            foreach (Guideline guideline in Guidelines)
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
            string date = System.DateTime.Now.ToString("dd-MM-yy-H-mm-ss");
            doc.Save(pathToChapterDocumentFolder + $@"\{date}_Guidelines.xml");
            if (WordDocGuidelineTools._WarningLogHasWarnings)
            {
                WordDocGuidelineTools._ErrorLog.Save(pathToChapterDocumentFolder + $@"\{date}_WarningLog.xml");
            }
            else {
                WordDocGuidelineTools._ErrorLog.Root.Add(new XElement("NoWarnings"));
                WordDocGuidelineTools._ErrorLog.Save(pathToChapterDocumentFolder + $@"\{date}_WarningLog.xml");
            }


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

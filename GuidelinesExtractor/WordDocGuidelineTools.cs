using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;


namespace GuidelinesExtractor
{
    public static class WordDocGuidelineTools
    {

        static public Word.Application _WordApp = OpenWordApp();
        static public Word.Document _ChapterWordDoc;

        static public XDocument _ErrorLog = new XDocument(new XElement("Warnings"));

        static public bool _WarningLogHasWarnings = false;

        public enum ExtractionMode
        {
            NoBookmarking = 0,
            BookmarkAllGuidelines = 1,
            BookmarkOnlyNewGuidelinesAndCheckForChangesOfPreviouslyBookmarkedGuidelines = 2
        }

        public static void OpenDocument(String docPath)
        {
            _ChapterWordDoc = _WordApp.Documents.Open(docPath);
        }

        public static Word.Application OpenWordApp()
        {
            Word.Application wordApp = new Word.Application();
            // Make Word visible (optional).
            wordApp.Visible = true;
            wordApp.Activate();

            return wordApp;
        }

        /// <summary>
        /// Extracts individual Guidelines from document based on given style
        /// </summary>
        public static List<Guideline> GetGuideLinesInDocument(string chapterWordFilePath, int chapterNumber, string guidelineTitleStyle, ExtractionMode extractionMode)
        {
            OpenDocument(chapterWordFilePath);

            //object guideLineStyle = GetDocumentGuideLineStyle(); //chapters are inconsistent with styling and fonts

            List<Guideline> guidelines = new List<Guideline>();

            Word.Range rng = _WordApp.ActiveDocument.Content;

            rng.Find.ClearFormatting();
            rng.Find.Text = "Guidelines";
            rng.Find.set_Style(guidelineTitleStyle);

            rng.Find.Execute();

            bool foundAGuidelineTable = false;
            while (rng.Find.Found)
            {
                // the range of rng will just be the word "Guidelines" which is in a table. So the rngTables
                //will just be one Table which is the table that the Guideline is in. 
                foreach (Word.Table guidelineTable in rng.Tables)
                {
                    GetGuidelinesInTable(ref guidelines, chapterNumber, guidelineTable, extractionMode);
                    // GetGuidelineFromTable(ref guidelines, guidelineTable);
                }
                foundAGuidelineTable = true;
                rng.Find.Execute();

            }

            if (!foundAGuidelineTable)// no guidelines found
            {
                LogErrors(chapterNumber, guidelineTitleStyle);
                return guidelines;
            }

            _WordApp.Documents.Close(SaveChanges: Word.WdSaveOptions.wdPromptToSaveChanges);
            return guidelines;

        }

        private static void LogErrors(int chapterNumber, string guidelineTitleStyle)
        {

            var newWarning = new XElement("warning");
            newWarning.Value = $"No guidelines were found in chapter \"{chapterNumber}\". Check if the Guideline Title Style is \"{guidelineTitleStyle}\". If not rerun using the Guideline Title style used in that file and the associated (latest created) output xml as the input to preserve changes that were made by this run.";
            _ErrorLog.Root.Add(newWarning);
            _WarningLogHasWarnings = true;

        }

        /// <summary>
        /// Attempts to discover the style use for the Guideline title in the document
        /// </summary>
        private static object GetDocumentGuideLineStyle()
        {
            Word.Range rng = _WordApp.ActiveDocument.Content;

            rng.Find.ClearFormatting();
            rng.Find.Text = "Guidelines^p";
            rng.Find.Font.Color = Word.WdColor.wdColorBlack;

            rng.Find.Execute();
            if (rng.Find.Found == true)
            {
                return rng.get_Style();
            }
            else
            {

                return null;
            }

        }

        /* public static void GetGuidelineFromTable(ref List<string> guidelineTable, Word.Table table)
         {

             for (int row = 1; row <= table.Rows.Count; row++)
             {
                 var cell = table.Cell(row, 1);
                 var text = cell.Range.Text;
                 if (text.Contains("Guidelines"))
                 {
                     guidelineTable.Add(text);
                 }
                 // text now contains the content of the cell.
             }

         }*/


        public static void GetGuidelinesInTable(ref List<Guideline> guidelines, int chapterNumber, Word.Table table, ExtractionMode extractionMode)
        {
            Word.Range individualGuidelineRange;
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                var cell = table.Cell(row, 1);
                var text = cell.Range.Text;
                if (text.Contains("Guidelines"))
                {

                    individualGuidelineRange = cell.Range;

                    ParseIndividualGuidelinesInTable(cell.Range, ref guidelines, chapterNumber, extractionMode);

                }

            }

        }

        /// <summary>
        /// Bookmarks each individal guideline in the Guideline table. e.g. Ch1_123asd8ad
        /// </summary>
        private static void ParseIndividualGuidelinesInTable(Range tableRange, ref List<Guideline> guidelines, int chapterNumber, ExtractionMode extractionMode)
        {

            string guidAsString;

            string tableText = tableRange.Text;
            MatchCollection guidelineMatches;

            if (!string.IsNullOrEmpty(tableText))
            {

                guidelineMatches = Regex.Matches(tableText, @"(([^\\](?<!\r))*(?=(\r)))"); //text followed by carriage return -> get each individual guideline (i.e. each  DO..., DONT... etc) in the the table
            }
            else { return; }

            object start = tableRange.Start;
            object end = tableRange.End;

            Word.Range individualGuidelineRange;

            foreach (Match guidelineMatch in guidelineMatches)
            {

                if (guidelineMatch.Value.StartsWith("Guideline") || string.IsNullOrWhiteSpace(guidelineMatch.Value)) continue;//skip the first line and whitespace matches

                string guidelineText = GetXmlCompatibleText(guidelineMatch.Value);

                individualGuidelineRange = _ChapterWordDoc.Range(start, end);
                string searchText = PrepareFindText(guidelineMatch.Value);

                individualGuidelineRange.Find.Text = searchText;
                individualGuidelineRange.Find.Execute();


                if (individualGuidelineRange.Find.Found)
                {
                    guidAsString = Guid.NewGuid().ToString("N");
                    string bookmarkKey = "NA";
                    

                    Guideline guideline;
                    switch (extractionMode)
                    {
                        case ExtractionMode.BookmarkAllGuidelines:
                            bookmarkKey = (_ChapterWordDoc.Bookmarks.Add(($"Ch{chapterNumber.ToString().PadLeft(2, '0')}_{guidAsString}").Substring(0, 12), individualGuidelineRange).Name);
                            guideline = new Guideline() { Key = bookmarkKey, Text = guidelineText };
                            GuidelinesFormatter.Guidelines.Add(guideline);
                            break;

                        case ExtractionMode.NoBookmarking:
                            bookmarkKey = "NA";
                            guideline = new Guideline() { Key = bookmarkKey, Text = guidelineText };
                            GuidelinesFormatter.Guidelines.Add(guideline);
                            break;

                        case ExtractionMode.BookmarkOnlyNewGuidelinesAndCheckForChangesOfPreviouslyBookmarkedGuidelines:
                            //check if text is already bookmarked and if text has changed.
                            GuidelineStatus guidelineStatus = checkBookmarkStatus(guidelineText, individualGuidelineRange, ref bookmarkKey);
                            if (guidelineStatus == GuidelineStatus.NotPreviouslyBookmarkedNewGuideline)
                            {
                                bookmarkKey = (_ChapterWordDoc.Bookmarks.Add(($"Ch{chapterNumber.ToString().PadLeft(2, '0')}_{guidAsString}").Substring(0, 12), individualGuidelineRange).Name);
                                guideline = new Guideline() { Key = bookmarkKey, Text = guidelineText };
                                GuidelinesFormatter.Guidelines.Add(guideline);
                            }
                            else if (guidelineStatus == GuidelineStatus.BookmarkedAndChanged)
                            {

                                Guideline existingGuideLine;
                                Guideline guidelineWithChangedText = new Guideline() { Key = bookmarkKey };
                                if (GuidelinesFormatter.Guidelines.TryGetValue(guidelineWithChangedText, out existingGuideLine))
                                {
                                    GuidelinesFormatter.Guidelines.Remove(guidelineWithChangedText); //remove old
                                    guideline = new Guideline() { Key = bookmarkKey, Text = guidelineText, Comments = existingGuideLine.Comments, Section = existingGuideLine.Section, Severity = existingGuideLine.Severity, Subsection = existingGuideLine.Subsection };//update new
                                    GuidelinesFormatter.Guidelines.Add(guideline);

                                }
                            }

                            break;
                    }

                    //guidelines.Add((bookmark, individualGuidelineRange.Text));

                    //GuidelinesFormatter.Guidelines.Add(guideline);

                    //guidelines.Add(guideline);

                }
            }

        }

        private static string GetXmlCompatibleText(string text)
        {
            if (!IsValidXmlString(text))
            {
                return RemoveInvalidXmlChars(text);
            }
            else return text;
        }

        /// <summary>
        /// Determines the state of a guideline. If a bookmark already exists for it than the ref bookmarkKey is set to its id.
        /// </summary>
        private static GuidelineStatus checkBookmarkStatus(string guideLineText, Range individualGuidelineRange, ref string bookmarkKey)
        {
            var bookmarksEnumerator = individualGuidelineRange.Bookmarks.GetEnumerator();
            int guidelineStatus = 0;
            if (!bookmarksEnumerator.MoveNext()) { return (GuidelineStatus)guidelineStatus; }
            else //get bookmark ID
            {
                guidelineStatus++;
                var bookmark = (Word.Bookmark)bookmarksEnumerator.Current;
                bookmarkKey = bookmark.Name;

                var currentDocumentGuideline = new Guideline();

                currentDocumentGuideline.Key = bookmarkKey;
                Guideline existingGuideLine;
                if (GuidelinesFormatter.Guidelines.TryGetValue(currentDocumentGuideline, out existingGuideLine))
                { //check if text has changed.

                    int distance = GuidelineTextCompare.Levenshtein(guideLineText, existingGuideLine.Text);

                    if (distance > 2)
                    {
                        guidelineStatus++;
                    }

                }
                //find that bookmark key's guideline text from the previous set of guidelines. compare to current text.
                //guidelineStatus++;

                return (GuidelineStatus)guidelineStatus;
                //return false;
            }

        }

        private enum GuidelineStatus
        {

            NotPreviouslyBookmarkedNewGuideline = 0,
            BookmarkedAndNoChange = 1,
            BookmarkedAndChanged = 2

        }

        static bool IsValidXmlString(string text)
        {
            try
            {
                XmlConvert.VerifyXmlChars(text);
                return true;
            }
            catch
            {
                return false;
            }
        }

        static string RemoveInvalidXmlChars(string text)
        {
            var validXmlChars = text.Where(ch => XmlConvert.IsXmlChar(ch)).ToArray();
            return new string(validXmlChars);
        }

        private static string PrepareFindText(string value) //check for special characters
        {
            if (value.Length > 255)
            {
                value = value.Substring(0, 254); //a word search only can be 255 characters
            }

            return value.Replace("^", "^^");


        }
    }
}

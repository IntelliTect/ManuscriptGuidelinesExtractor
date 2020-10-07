using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;


namespace GuidelinesExtractor
{
    public static class GuideLineTools
    {

        static public Word.Application _WordApp = OpenWordApp();
        static public Word.Document _ChapterWordDoc;

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



        public static List<(string, string)> GetGuideLinesInDocument(string chapterWordFilePath, int chapterNumber, string guidelineTitleStyle)
        {
            OpenDocument(chapterWordFilePath);

            //object guideLineStyle = GetDocumentGuideLineStyle(); //chapters are inconsistent with styling and fonts

            List<Guideline> guidelines = new List<Guideline>();



            Word.Range rng = _WordApp.ActiveDocument.Content;

            rng.Find.ClearFormatting();
            rng.Find.Text = "Guidelines";
            rng.Find.set_Style(guidelineTitleStyle);

            rng.Find.Execute();

            while (rng.Find.Found)
            {
                // the range of rng will just be the word "Guidelines" which is in a table. So the rngTables
                //will just be one Table which is the table that the Guideline is in. 
                foreach (Word.Table guidelineTable in rng.Tables)
                {
                    GuidelineBookmarking(ref guidelines, chapterNumber, guidelineTable);
                    // GetGuidelineFromTable(ref guidelines, guidelineTable);
                }

                rng.Find.Execute();

            }



            _WordApp.Documents.Close(SaveChanges: Word.WdSaveOptions.wdPromptToSaveChanges);
            return guidelines;

        }

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

        public static void GetGuidelineFromTable(ref List<string> guidelineTable, Word.Table table)
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

        }


        public static void GuidelineBookmarking(ref List<Guideline> guidelines, int chapterNumber, Word.Table table)
        {
            Word.Range individualGuidelineRange;
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                var cell = table.Cell(row, 1);
                var text = cell.Range.Text;
                if (text.Contains("Guidelines"))
                {

                    individualGuidelineRange = cell.Range;

                    BookmarkGuidelinesInTable(cell.Range, ref guidelines, chapterNumber);

                }

            }

        }

        private static void BookmarkGuidelinesInTable(Range tableRange, ref List<Guideline> guidelines, int chapterNumber)
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

                individualGuidelineRange = _ChapterWordDoc.Range(start, end);
                string searchText = PrepareFindText(guidelineMatch.Value);

                individualGuidelineRange.Find.Text = searchText;
                individualGuidelineRange.Find.Execute();

               

                if (individualGuidelineRange.Find.Found)
                {
                    guidAsString = Guid.NewGuid().ToString("N");
                    string bookmark = (_ChapterWordDoc.Bookmarks.Add(($"Ch{chapterNumber.ToString().PadLeft(2,'0')}_{guidAsString}").Substring(0, 12), individualGuidelineRange).Name);
                    //guidelines.Add((bookmark, individualGuidelineRange.Text));
                    Guideline guideline = new Guideline() {Key=bookmark,Text=individualGuidelineRange.Text};
                    guidelines.Add(guideline);

                }

            }

        }

        private static string PrepareFindText(string value) //check for special characters
        {
            if (value.Length > 255) { 
            value= value.Substring(0, 254); //a word search only can be 255 characters
            }

            return value.Replace("^", "^^");

            
        }
    }
}
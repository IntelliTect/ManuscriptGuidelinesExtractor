using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
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



        public static List<string> GetGuideLinesInDocument(string chapterWordFilePath, string guidelineTitleStyle)
        {
            OpenDocument(chapterWordFilePath);

            //object guideLineStyle = GetDocumentGuideLineStyle(); //chapters are inconsistent with styling and fonts

            List<string> guidelines = new List<string>();
           

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
                        GuidelineBookmarking(ref guidelines, guidelineTable);
                       // GetGuidelineFromTable(ref guidelines, guidelineTable);
                    }

                    rng.Find.Execute();

                }
            

           
           // _WordApp.Documents.Close(SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges);
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

        public static void GetGuidelineFromTable(ref List<string> guidelines, Word.Table table)
        {

            for (int row = 1; row <= table.Rows.Count; row++)
            {
                var cell = table.Cell(row, 1);
                var text = cell.Range.Text;
                if (text.Contains("Guidelines"))
                {
                    guidelines.Add(text);
                  

                }
                // text now contains the content of the cell.
            }

        }


        public static void GuidelineBookmarking(ref List<string> guidelines, Word.Table table)
        {
            Word.Range individualGuidelineRange;
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                var cell = table.Cell(row, 1);
                var text = cell.Range.Text;
                if (text.Contains("Guidelines"))
                {

                    individualGuidelineRange = cell.Range;

                    BookmarkGuidelinesInTable(cell.Range, ref guidelines);


                    //guidelines.Add(text);


                }
                // text now contains the content of the cell.
            }

        }

        private static void BookmarkGuidelinesInTable(Range tableRange, ref List<string> guidelines)
        {
            Word.Range individualGuidelineRange = tableRange;

            List<string> bookmarkGuids = new List<string>();
            string guidAsString;

            individualGuidelineRange.Find.ClearFormatting();
            individualGuidelineRange.Find.Text = "^p";

            // 
            object start = individualGuidelineRange.Start;
            object end = individualGuidelineRange.End;
            Word.Range startingRange = _ChapterWordDoc.Range(start, end);
            individualGuidelineRange.Find.Execute();

            Range guidelineEndingWithParagraphRange = _ChapterWordDoc.Range(start, end);
            while (individualGuidelineRange.Find.Found)
            {
                if (individualGuidelineRange.Start >= (int)end) //no longer in table text
                {
                    break;
                }
                guidelineEndingWithParagraphRange.Start = startingRange.Start; //from the start of text
                guidelineEndingWithParagraphRange.End = individualGuidelineRange.Start; // to the first paragraph mark

                if (!guidelineEndingWithParagraphRange.Text.StartsWith("Guideline"))//skip the first line ending with a paragraph mark
                {
                    guidAsString = Guid.NewGuid().ToString("N");
                    guidelines.Add(guidelineEndingWithParagraphRange.Text); //bookmark this range
                    
                    Range guidelineTextRangeWithNoParagraphCharacters = _ChapterWordDoc.Range(guidelineEndingWithParagraphRange.Start + 1, guidelineEndingWithParagraphRange.End); //increment the start to ignore the paragraph special character
                    
                    bookmarkGuids.Add(_ChapterWordDoc.Bookmarks.Add(($"a{guidAsString}").Substring(0, 12), guidelineEndingWithParagraphRange).Name);
                }

                startingRange.Start = individualGuidelineRange.Start;


                individualGuidelineRange.Find.Execute();

            }
            guidAsString = Guid.NewGuid().ToString("N");
            startingRange.End = (int)(end)-1; //we can then check if no paragraph was found we can just get all the text from the startRange.Start to the original end of the table.
            startingRange.Start = startingRange.Start + 1;
            guidelines.Add(startingRange.Text); //bookmark this range
            _ChapterWordDoc.Bookmarks.Add(($"a{guidAsString}").Substring(0,12), startingRange);




        }
    }
}

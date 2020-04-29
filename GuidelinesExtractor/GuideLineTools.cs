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



        public static List<string> GetGuideLinesInDocument(string chapterWordFilePath)
        {
            OpenDocument(chapterWordFilePath);

            object guideLineStyle = GetDocumentGuideLineStyle(); //chapters are inconsistent with styling and fonts

            List<string> guidelines = new List<string>();
            if (guideLineStyle != null) //atleast one guideline was found
            {

                Word.Range rng = _WordApp.ActiveDocument.Content;

                rng.Find.ClearFormatting();
                rng.Find.Text = "Guidelines";
                rng.Find.set_Style(guideLineStyle);

                rng.Find.Execute();

                while (rng.Find.Found)
                {
                    // the range of rng will just be the word "Guidelines" which is in a table. So the rngTables
                    //will just be one Table which is the table that the Guideline is in. 
                    foreach (Word.Table guidelineTable in rng.Tables)
                    {
                        GetGuidelineFromTable(ref guidelines, guidelineTable);
                    }

                    rng.Find.Execute();

                }
            }

           
            _WordApp.Documents.Close(SaveChanges: Word.WdSaveOptions.wdDoNotSaveChanges);
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
                    //TranslateWordStylingToMarkDown(cell.Range);

                }
                // text now contains the content of the cell.
            }

        }

        private static string TranslateWordStylingToMarkDown(Range range)
        {
            throw new NotImplementedException();
        }
    }
}

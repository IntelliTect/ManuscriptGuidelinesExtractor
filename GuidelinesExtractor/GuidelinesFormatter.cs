using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace GuidelinesExtractor
{
    public class GuidelinesFormatter
    {

        /*        public static void AllGuidelinesToMarkDown(string pathToChapterDocumentFolder, bool verbose, string guidelineTitleStyle)
                {

                    List<string> allDocs = FileManager.GetAllFilesAtPath(pathToChapterDocumentFolder, searchPattern: "Michaelis_Ch??.docx")
                        .OrderBy(x => x).ToList();

                    List<string> AllGuidelines = new List<string>();

                    StreamWriter markdownFileWriter = File.CreateText(pathToChapterDocumentFolder + @"/Guidelines.md");

                    string h3 = "###";
                    string h1 = "#";

                    int chapterNumber = 0;
                    foreach (string chapterDocxPath in allDocs)
                    {
                        List<string> currentChapterGuidelines = GuideLineTools.GetGuideLinesInDocument(chapterDocxPath, guidelineTitleStyle);

                        if (currentChapterGuidelines.Count > 0)
                        {
                            chapterNumber = GetChapterNumber(chapterDocxPath);
                            markdownFileWriter.WriteLine($"{h1} Chapter {chapterNumber}");


                            char[] splitChars = { '\r' };

                            foreach (string guideline in currentChapterGuidelines)
                            {
                                string[] guidelineParts = guideline.Split(splitChars);

                                markdownFileWriter.WriteLine($"{h3} {guidelineParts[0]}"); // "Guidelines"

                                for (int i = 1; i < guidelineParts.Length; i++)
                                {
                                    if (guidelineParts[i].Length > 2) //ignore words \a (not sure if all guidelines have this format so just using a length check )
                                    markdownFileWriter.WriteLine("- " + guidelineParts[i]);
                                }

                            }


                            if (verbose)
                            {
                                //print all guidelines
                                Console.WriteLine($"Chapter {chapterNumber}");
                                foreach (string guideline in currentChapterGuidelines)
                                {
                                    Console.WriteLine(guideline);
                                }
                            }
                        }

                    }


                    markdownFileWriter.Close();
                    GuideLineTools._WordApp.Quit();

                }
                */



        public static void AllGuidelinesToCSVWithBookmarks(string pathToChapterDocumentFolder, bool verbose, string guidelineTitleStyle)
        {

            List<string> allDocs = FileManager.GetAllFilesAtPath(pathToChapterDocumentFolder, searchPattern: "Michaelis_Ch??.docx")
                .OrderBy(x => x).ToList();

            List<string> AllGuidelines = new List<string>();

            StreamWriter csvFileWriter = File.CreateText(pathToChapterDocumentFolder + @"/Guidelines.csv");

            string h3 = "###";
            string h1 = "#";

            int chapterNumber = 0;
            List<(string, string)> currentChapterGuidelines;

            //before your loop
            var csv = new System.Text.StringBuilder();

            
            foreach (string chapterDocxPath in allDocs)
            {
               chapterNumber = GetChapterNumber(chapterDocxPath);
               currentChapterGuidelines= (GuideLineTools.GetGuideLinesInDocument(chapterDocxPath, chapterNumber, guidelineTitleStyle));

                foreach ((string, string) bookmarkAndGuideline in currentChapterGuidelines) {
                    csvFileWriter.WriteLine($"{bookmarkAndGuideline.Item1},\"{bookmarkAndGuideline.Item2}\"");
                }

            }

           
            csvFileWriter.Close();
            GuideLineTools._WordApp.Quit();

        }




        public static void AllGuidelinesToXMLWithBookmarks(string pathToChapterDocumentFolder, bool verbose, string guidelineTitleStyle)
        {

            List<string> allDocs = FileManager.GetAllFilesAtPath(pathToChapterDocumentFolder, searchPattern: "Michaelis_Ch??.docx")
                .OrderBy(x => x).ToList();

            List<string> AllGuidelines = new List<string>();

           

            string h3 = "###";
            string h1 = "#";

            int chapterNumber = 0;
            List<(string, string)> currentChapterGuidelines = new List<(string, string)>();

            //before your loop
            var csv = new System.Text.StringBuilder();


            foreach (string chapterDocxPath in allDocs)
            {
                chapterNumber = GetChapterNumber(chapterDocxPath);
                //instead of the following line using GetGuideLinesInDocument. a new method called GetUniqueGuideLinesInDocument(exiting guidelines stored in xml) 
                //can be used to just find new guidelines (by determining if the guideline has a bookmark on it(note bookmarks only cover 255 characters)) and append them to the xml file 
                currentChapterGuidelines.AddRange(GuideLineTools.GetGuideLinesInDocument(chapterDocxPath, chapterNumber, guidelineTitleStyle));

                WriteXML(currentChapterGuidelines, pathToChapterDocumentFolder);

            }

            GuideLineTools._WordApp.Quit();

        }

        private static void WriteXML(List<(string, string)> currentChapterGuidelines, string pathToChapterDocumentFolder)
        {

            StreamWriter xmlFileWriter = File.CreateText(pathToChapterDocumentFolder + @"/Guidelines.csv");

            xmlFileWriter.WriteLine("<?xml version=\"1.0\" encoding=\"utf - 8\"?>\n< root >\n");
            // example format
            // <guideline key="Ch01_fa67753" severity="DO" section="Naming" subsection="Variables and fields">DO favor clarity over brevity when naming identifiers.</guideline>
            foreach ((string, string) bookmarkAndGuideline in currentChapterGuidelines)
            {
                xmlFileWriter.WriteLine($"\t<guideline key=\"{bookmarkAndGuideline.Item1} severity=\"\" section=\"\" subsection=\"\">\"{bookmarkAndGuideline.Item2}\"</guideline>");
            }

            xmlFileWriter.WriteLine("<\\root >");

            xmlFileWriter.Close();
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

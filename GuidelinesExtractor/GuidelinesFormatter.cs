using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace GuidelinesExtractor
{
    public class GuidelinesFormatter
    {

        public static void AllGuidelinesToMarkDown(string pathToChapterDocumentFolder, bool verbose)
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
                List<string> currentChapterGuidelines = GuideLineTools.GetGuideLinesInDocument(chapterDocxPath);

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

using CommandLine;
using System;
using System.IO;
using System.Linq;

namespace GuidelinesExtractor
{
    public class Program
    {
        private const string IntelliTect =
@" _____       _       _ _ _ _______        _   
|_   _|     | |     | | (_)__   __|      | |  
  | |  _ __ | |_ ___| | |_   | | ___  ___| |_ 
  | | | '_ \| __/ _ \ | | |  | |/ _ \/ __| __|
 _| |_| | | | ||  __/ | | |  | |  __/ (__| |_ 
|_____|_| |_|\__\___|_|_|_|  |_|\___|\___|\__|";

        private const string InteractivePromptPrefix = "INTL {0} ({1})>";


        public class Options
        {
            [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
            public bool Verbose { get; set; }

            [Option('a', "all", Required = false, HelpText = "All chapters to markdown.")]
            public bool All { get; set; }

            [Option('p', "path", Required = true, HelpText = "path to folder with chapters")]
            public string PathToFolder { get; set; }
        }

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       if (o.Verbose)
                       {
                           Console.WriteLine($"Verbose output enabled. Current Arguments: -v {o.Verbose}");
                           Console.WriteLine("Quick Start Example! App is in Verbose mode!");
                       }
                       else
                       {
                           Console.WriteLine($"Current Arguments: -v {o.Verbose}");
                           Console.WriteLine("Quick Start Example!");
                       }
                       if (o.All && String.IsNullOrEmpty(o.PathToFolder)) {
                           Console.WriteLine("Doing all Guidelines");
                           Run(path: o.PathToFolder, verbose: o.Verbose);
                       }
                   });
        }


        public static void Run(string path = "", Modes mode = Modes.GetAllGuidelines,
            bool verbose = false,
            bool preview = false,
            bool byFolder = false,
            bool chapterOnly = false)
        {
            /*var colorList = new List<ConsoleColor>{ConsoleColor.Blue, ConsoleColor.Green, ConsoleColor.Yellow, 
                ConsoleColor.DarkCyan, ConsoleColor.DarkRed, ConsoleColor.Cyan};

            var intelliTectSplit = IntelliTect.Split(new[] {Environment.NewLine}, StringSplitOptions.None);
            for (var index = 0; index < intelliTectSplit.Length; index++)
            {
                string line = intelliTectSplit[index];
                ConsoleColor prevColor = Console.ForegroundColor;
                Console.ForegroundColor = colorList[index];
                Console.WriteLine(line);
                Console.ForegroundColor = prevColor;
            }*/

            Console.WriteLine(IntelliTect);

            if (preview)
            {
                Console.WriteLine("Preview mode. Actions will not be taken");
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                ConsoleColor foregroundColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Yellow;
                path = Environment.CurrentDirectory;
                Console.WriteLine($"Path not specified. Using working directory: {path}");
                Console.ForegroundColor = foregroundColor;
            }

            switch (mode)
            {
                case Modes.GetAllGuidelines:
                    Console.WriteLine($"Getting all Guidelines from the documents found in: {path}");
                    GuidelinesFormatter.AllGuidelinesToMarkDown(path, verbose);
                    break;
              /*  case Modes.TestGeneration:
                    var generatedTests
                        = ListingManager.GenerateUnitTests(path, TestGeneration_Interactive, true);
                    if (verbose)
                    {
                        Console.WriteLine($"{generatedTests.Count} tests generated");
                    }
                    break;
                case Modes.ScanForMismatchedListings:
                    var extraListings = ListingManager.GetAllExtraListings(path).OrderBy(x => x);

                    Console.WriteLine("---Extra Listings---");
                    foreach (string extraListing in extraListings)
                    {
                        Console.WriteLine(extraListing);
                    }
                    break;*/
                default:
                    Console.WriteLine($"Mode ({mode}) does not exist. Exiting");
                    break;
            }
        }

        private static bool TestGeneration_Interactive(string missingTest)
        {
            InteractiveConsoleWrite("Choose an option", "d - delete, q - quit, enter - continue");
            string input = Console.ReadLine();

            switch (input)
            {
                case "d":
                    Console.WriteLine("Deleting test");
                    File.Delete(missingTest);
                    break;
                case "q":
                    Console.WriteLine("Quitting");
                    return false;
            }

            return true;
        }

        private static void InteractiveConsoleWrite(string toWrite, string userOptions)
        {
            Console.Write(InteractivePromptPrefix, toWrite, userOptions);
        }
    }
}

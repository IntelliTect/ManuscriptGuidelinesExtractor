# Manuscript Guidelines Extractor (Retired)

## Use

If you want to extract an XML file from word documents:

1. Locate test called RUN_Program.
2. Copy all word docs to get guidelines to the WordDocs folder.
3. Update `string wordDocFolder` with the absolute path to the WordDocs Folder.
4. Run the test. The outputted XML should have the compilation date and time in the name and is located in the WordDocs folder.

## Expected Functionality (Not working completely.)

Theoretically, this program should run via console which it doesn't currently and also the test that is used to run it currently should in GuidelinesFormatter in its third argument take `WordDocGuidelineTools.ExtractionMode.BookmarkOnlyNewGuidelinesAndCheckForChangesOfPreviouslyBookmarkedGuidelines` and `pathToExistingGuidelines` in the 4th argument, and be able to compare against an existing guidelines XML file and add guidelines that are new in the WordDocs. This is done by comparing the bookmarks already in the word doc and checking if there are any new guidelines that are not yet bookmarked and consider this a new guideline.
However, there seems to be errors and problems with this.

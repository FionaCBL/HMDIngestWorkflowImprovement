# HMD Ingest workflow tool 
This tool, written in [C#](https://en.wikipedia.org/wiki/C_Sharp_(programming_language)), provides functionality for reading values from the HMD sharepoint site and automating the ingest workflow process by performing various operations on the values returned from sharepoint. The tool provides outputs in various formats, in particular by writing to text and csv files but also by writing directly to the sharepoint site.

# Table of Contents
1. [Structure of the tool](#Structure-of-the-tool)
2. [Using the tool](#Using-the-tool)


## Structure of the tool
The tool is controlled with the [Program.cs](./Program.cs) file, with different functionality split across the other `.cs` files in the repository: 

- SharepointTools.cs
    - Functions to interact with the HMD sharepoint site
    - Verifies site exists, loads from the digitisation workflow sharepoint list and prints specified fields for each item (shelfmark)
    - These functions provide an input for the other functions defined in the `.cs` files listed below
- DirectorySearchTools.cs
    - Contains functions to list the file contents of folder paths passed as an argument, including option to search recursively
- InputOrderSpreadsheetTools.cs
    - Contains functions to generate input order CSV file for each shelfmark provided
- PathFinder.cs	(*UNDER CONSTRUCTION* - To be reviewed)
    - Functions to derive full path for a given shelfmark when given the \\\ad\collections\ version of the source folder path
- TextFileOutput.cs
    - Functions to output text files after running some analysis on items in sharepoint, mainly used for development purposes so far

## Using the tool





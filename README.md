# HMD Ingest workflow tool 
This tool, written in [C#](https://en.wikipedia.org/wiki/C_Sharp_(programming_language)), provides functionality for reading values from the HMD sharepoint site and automating the ingest workflow process by performing various operations on the values returned from sharepoint. The tool provides outputs in various formats, in particular by writing to text and csv files but also by writing directly to the sharepoint site.


## Users
The tool is supplied as a Windows executable file which can be run with no prior installation.

To run, double-click the executable located in   
```G:\Heritage Made Digital\05\Projects\Workflow\Validation of Ingest Workflow\SoftwareTesting```

You will be presented with a series of yes/no options.

### Environment

Two environments exist within the software tool - `test` and `prod`. Using `test` sets up a test-based environment, only looks at the test sharepoint site and writes output files to your desktop. Using `prod` sets up the live production environment and uses the live HMD sharepoint.  
**Currently writing-to-sharepoint functionality is not working, so both environments write to the user's desktop** 

You can also select a `project` to run over. The default project is environment-dependent, but the program will tell you what it is currently set to. To change it, type `yes` (then `Enter`) and type the name of the project to analyse. At the moment 
this is required to match the value in SharePoint exactly.

### Checks
As far as possible, individual checks are separated and the user is free to choose which checks to run over. 

Three broad checks are available:
- Shelfmark source folder checks
This checks that the value of `Source Folder` given for an item in SharePoint is a valid path on the BL network. If not, a variable is written to SharePoint. **Sharepoint writing not yet available**

- Run shelfmark protected character check
Checks a shelfmark as entered in SharePoint for protected characters that will cause errors further down the workflow. Writes to SharePoint if errors are found. **Sharepoint writing not yet available**

- Generates image order CSV and perform ALTO XML checks. Writes to SharePoint if errors are found. 
**Sharepoint writing not yet available**

### For testing the week of 18th June 2020
The image order CSV is the most complete check and offers the most useful output so it would be great if you could test this! Image order CSVs should appear in your Desktop within the HMDSharepoint_ImgOrderCSVs folder regardless of which environment you select. Feel free to test things with the `test` or `prod` environments, and use any project you like. The `prod` environment was only included because I started having issues accessing the test SharePoint on 17/06/2020.



## For Developers
### Structure of the tool
The tool is controlled with the [Program.cs](./Program.cs) file, with different functionality split across the other `.cs` files in the repository: 

- SharepointTools.cs
    - Functions to interact with the HMD sharepoint site
    - Verifies site exists, loads from the digitisation workflow sharepoint list and prints specified fields for each item (shelfmark)
    - These functions provide an input for the other functions defined in the `.cs` files listed below
- DirectorySearchTools.cs
    - Contains functions to list the file contents of folder paths passed as an argument, including option to search recursively
- InputOrderSpreadsheetTools.cs
    - Contains functions to generate input order CSV file for each shelfmark provided
- TextFileOutput.cs
    - Functions to output text files after running some analysis on items in sharepoint, mainly used for development purposes so far






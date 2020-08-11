using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelDataReader;



namespace HMDSharepointChecker
{
    class Program
    {
        static bool RunInitialTests(string spURL)
        {
            // Test sharepoint site exists
            try
            {

                Assert.IsTrue(SharepointTools.SharepointSiteExists(spURL));

                // Tests that you can retrieve a tile from the sharepoint site
                var spTitle = SharepointTools.GetSharepointTitle(spURL);
                Assert.IsNotNull(spTitle);
                Console.WriteLine("Sharepoint site title: {0}", spTitle);
                var spListNames = SharepointTools.GetSharePointLists(spURL);
                Assert.IsTrue(spListNames.Count != 0);

                // test a single XML version number
                bool goodXMLVersion = false;
                String testXMLPath = @"\\v8l-lon2\DATA\Heritage Made Digital\05 Projects\Workflow\Validation of Ingest Workflow\Example OCR files\IOR!P!5644_Jun_1899_nos_80-92_001.xml";
                String XMLVersionNumber = InputOrderSpreadsheetTools.GetXMLVersionNumber(testXMLPath);
                float XMLVNum = float.Parse(XMLVersionNumber, System.Globalization.CultureInfo.InvariantCulture);
                if (XMLVNum <= 2.0)
                {
                    goodXMLVersion = true;
                }
                Assert.IsTrue(goodXMLVersion);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in RunInitialTests: {0}", ex);
                return false;
            }

        }
        static Dictionary<string,string> InitialSetup(String env)
        {
            string spURL = "";
            string project = "";
            if (env.Equals("prod"))
            {
                spURL = "http://hmd.sharepoint.ad.bl.uk";
                project = "Zoroastrian Manuscripts";
            }

            else if (env.Equals("test"))
            {
                spURL = "http://v12t-sp13wfe1:88/";
                //project = "Zoroastrian Manuscripts";
                project = "HMD Portfolio - 16th Century English Manuscripts";
            }
            else
            {
                Console.WriteLine("You must set an environment variable.");
                return null;
            }

            Dictionary<string, string> dict = new Dictionary<string, string>{{ "spURL", spURL },{ "project", project }};
            return dict;
        }
        static void Main(string[] args)
        {
            bool runProgram = true;

            while (runProgram)
            {

                var startTime = DateTime.Now;
                bool debug = false;


                bool useIngestSpreadsheet = false;
                String restartProgram = String.Empty;

                // Start of sharepoint-specific items
                Console.WriteLine("This program can use the sharepoint site or a user-provided ingest spreadsheet.\nUse ingest spreadsheet?");
                String useInputIngestSpreadsheet = Console.ReadLine();
                if (useInputIngestSpreadsheet.ToLower().Equals("yes") || useInputIngestSpreadsheet.ToLower().Equals("y"))
                {
                    useIngestSpreadsheet = true;
                    Console.WriteLine("Using ingest spreadsheet...");
                }
                else
                {
                    Console.WriteLine("Using sharepoint...");
                }
                if (!useIngestSpreadsheet)
                {

                    // Set the environment (use 'test' or 'prod')
                    var env = "test";



                    Console.WriteLine("You are currently running {0}. Switch to prod? yes/no", env);
                    String inputEnv = Console.ReadLine();
                    if (inputEnv.ToLower().Equals("yes") || inputEnv.ToLower().Equals("y"))
                    {
                        env = "prod";
                    }


                    // do initial setup for sharepoint jobs
                    Dictionary<String, String> variablesDictionary = InitialSetup(env);
                    Assert.IsNotNull(variablesDictionary); // Check we set the variables properly
                    var spURL = variablesDictionary["spURL"];
                    var project = variablesDictionary["project"];
                    List<String> inputVariable = new List<string>();
                    Assert.IsTrue(RunInitialTests(spURL));

                    Console.WriteLine("This program searches sharepoint using either the 'Project' field or the 'Shelfmark' field.\nShelfmarks can be used individually or provided in a shelfmarks.txt file in this folder, with one shelfmark per line.\nUse shelfmarks? (yes/no)");
                    String useShelfmarksYN = Console.ReadLine();
                    if (useShelfmarksYN.ToLower().Equals("yes") || useShelfmarksYN.ToLower().Equals("y"))
                    {
                        Console.WriteLine("Use list of shelfmarks? (yes/no)");
                        String inputShelfmarkList = Console.ReadLine();
                        if (inputShelfmarkList.ToLower().Equals("yes") || inputShelfmarkList.ToLower().Equals("y"))
                        {
                            Console.WriteLine("Searching for 'shelfmarks.txt' file in current folder...");
                            var currentDir = Directory.GetCurrentDirectory();
                            var shelfmark_list_path = currentDir + "\\shelfmarks.txt";

                            if (System.IO.File.Exists(shelfmark_list_path))
                            {
                                Console.WriteLine("Found shelfmarks.txt file in current folder!");
                                foreach (var line in System.IO.File.ReadLines(shelfmark_list_path))
                                {
                                    inputVariable.Add(@"<FieldRef Name ='Title'/><Value Type = 'Text'>" + line + @"</Value>");
                                }


                            }
                            else
                            {
                                Console.WriteLine("Could not find file 'shelfmarks.txt' in current folder. Using single shelfmark...");
                                Console.WriteLine("Enter a shelfmark name (must match Sharepoint site)");
                                String inputSingleShelfmark = Console.ReadLine();
                                if (inputSingleShelfmark.Length > 0)
                                {
                                    inputVariable.Add(@"<FieldRef Name ='Title'/><Value Type = 'Text'>" + inputSingleShelfmark + @"</Value>");
                                }
                                else // kick it back and use defaults if nothing provided
                                {
                                    Console.WriteLine("No shelfmark provided! Using default project");
                                    inputVariable.Add(@"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>");
                                }
                            }
                        }
                        else
                        {

                            Console.WriteLine("Enter a shelfmark name (must match Sharepoint site)");
                            String inputSingleShelfmark = Console.ReadLine();
                            if (inputSingleShelfmark.Length > 0)
                            {
                                inputVariable.Add(@"<FieldRef Name ='Title'/><Value Type = 'Text'>" + inputSingleShelfmark + @"</Value>");
                            }
                            else // kick it back and use defaults if nothing provided
                            {
                                Console.WriteLine("No shelfmark provided! Using default project");
                                inputVariable.Add(@"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>");
                            }
                        }
                    }
                    else
                    {

                        Console.WriteLine("Project is currently set to {0}, change this? (yes/no)", project);
                        String inputProjectYN = Console.ReadLine();
                        if (inputProjectYN.ToLower().Equals("yes") || inputProjectYN.ToLower().Equals("y"))
                        {
                            Console.WriteLine("Type a project name (must match sharepoint record)", project);
                            String inputProject = Console.ReadLine();
                            if (inputProject.Length > 0)
                            {
                                inputVariable.Add(@"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + inputProject + @"</Value>");

                            }
                            else
                            {
                                Console.WriteLine("No project provided! Using default project.");
                                inputVariable.Add(@"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>");

                            }
                        }
                        else
                        {
                            Console.WriteLine("Using default project {0}", project);
                            inputVariable.Add(@"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>");

                        }
                    }

                    // Check we assigned something to the input variable
                    Assert.AreNotEqual(0, inputVariable.Count);

                    //===================== Checks to run ====================
                    bool reportShelfmarkCheckStatus = false;
                    bool runShelfmarkCharacterChecks = false;
                    bool runImageOrderGenerationXMLChecks = false;
                    bool queryMetadata = false;


                    Console.WriteLine("Run shelfmark source folder checks? yes/no");
                    String inputSourceFolderCheck = Console.ReadLine();
                    if (inputSourceFolderCheck.ToLower().Equals("yes") || inputSourceFolderCheck.ToLower().Equals("y"))
                    {
                        reportShelfmarkCheckStatus = true;
                    }
                    Console.WriteLine("Run shelfmark protected character check? yes/no");
                    String inputSMCharCheck = Console.ReadLine();
                    if (inputSMCharCheck.ToLower().Equals("yes") || inputSMCharCheck.ToLower().Equals("y"))
                    {
                        runShelfmarkCharacterChecks = true;
                    }
                    Console.WriteLine("Generate image order CSV and perform ALTO XML checks? yes/no");
                    String inputImageOrderGen = Console.ReadLine();
                    if (inputImageOrderGen.ToLower().Equals("yes") || inputImageOrderGen.ToLower().Equals("y"))
                    {
                        runImageOrderGenerationXMLChecks = true;
                    }
                    Console.WriteLine("Run query against descriptive metadata APIs? yes/no");
                    String inputMDAPIQuery = Console.ReadLine();
                    if (inputMDAPIQuery.ToLower().Equals("yes") || inputMDAPIQuery.ToLower().Equals("y"))
                    {
                        queryMetadata = true;
                    }
                    // =======================================================




                    // Get the 'Digitisation Workflow' list fields and print them out:
                    if (debug)
                    {
                        var DigitisationWorkflowTitles = SharepointTools.GetSharePointListFields(spURL, "Digitisation Workflow");
                        Assert.IsNotNull(DigitisationWorkflowTitles.Count);
                    }

                    // ============== Preliminary stuff - has to be run every time ================
                    // This is the "grab all the info from sharepoint" part of things

                    // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list

                    try
                    {
                        for (int counter = 0; counter < inputVariable.Count; ++counter)
                        {
                            var inputVar = inputVariable[counter];
                            Console.WriteLine("Processing input {0} of {1}", counter + 1, inputVariable.Count);


                            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListFieldContents(spURL, "Digitisation Workflow", env, inputVar);
                            if (DigitisationWorkflow_ID_Title_SourceFolders.Count < 1)
                            {
                                continue; // goes on to next input...
                            }

                            //  Check source folders - requires the above two lines to work
                            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
                            Assert.IsNotNull(SourceFolderStatus.Count);
                            if (debug)
                            {
                                Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus, "sourceFolderStatus"));
                            }

                            // =============================================================================

                            // ================= Optional checks - some take ~10 mins to read all of sharepoint, so don't run every time =============

                            // Validates that the value of the 'Source Folder' field exists as a valid path on the network
                            // Attempts to write the results to a sharepoint column for each row (== shelfmark)
                            if (reportShelfmarkCheckStatus)
                            {
                                Console.WriteLine("Writing source folder validation status to Sharepoint...");
                                String SharePointSourceFolderCheck = "SourceFolderValid";
                                Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointSourceFolderCheck));
                                Assert.IsTrue(SharepointTools.ReportSourceFolderStatus(spURL, "Digitisation Workflow", SharePointSourceFolderCheck, SourceFolderStatus));
                            }

                            // Checks for protected characters in Shelfmark names 
                            // Attempts to write the results to sharepoint column
                            if (runShelfmarkCharacterChecks)
                            {
                                Console.WriteLine("=======================================\nChecking shelfmarks for protected characters\n=======================================");
                                // Add in Shelfmark protected chars check here
                                String SharePointColumnShelfmarkCheck = "ShelfmarkCheck";
                                Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck));
                                List<HMDObject> badShelfmarks = SharepointTools.BadShelfmarkNames(SourceFolderStatus);
                                String shelfmarkCharacterStatus = "";

                                var thisItem = 1;

                                foreach (var item in badShelfmarks)
                                {
                                    if (badShelfmarks.Count > 20)
                                    {
                                        if (thisItem % 10 == 0)
                                        {
                                            Console.WriteLine("{0}/{1}", thisItem, badShelfmarks.Count);
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}/{1}", thisItem, badShelfmarks.Count);

                                    }
                                    thisItem += 1;
                                    var ID = item.ID;
                                    String SM = item.Shelfmark;
                                    if (item.BadShelfmark)
                                    {
                                        shelfmarkCharacterStatus = "Protected character(s) found";
                                    }
                                    else
                                    {
                                        shelfmarkCharacterStatus = "Valid";

                                    }
                                    Assert.IsTrue(SharepointTools.WriteToSharepointColumnByID(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck, SM, ID, shelfmarkCharacterStatus));
                                }


                            }

                            // Get the labels (image order, image type etc) for all shelfmarks passed into this function
                            // Write these out in an image order csv 


                            // Also runs ALTO XML checks - do they exist with the same names as TIFs and in the same number as TIFs?
                            // Are they version 2.0 or older?
                            if (runImageOrderGenerationXMLChecks)
                            {
                                bool writeToSharepoint = true;
                                bool addColumns = false;
                                if (addColumns)
                                {
                                    // Add columns for XML checking
                                    String SharePointColumnXMLCheck = "ALTOXMLCheck";
                                    Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnXMLCheck));

                                }
                                Console.WriteLine("=======================================\nQuerying metadata APIs to retrieve child shelfmark information...\n=======================================");
                                var iamsRecords = LibraryAPIs.queryMetadataAPIs(spURL, "Digitisation Workflow", SourceFolderStatus,writeToSharepoint);

                                var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus, env, spURL, "Digitisation Workflow", iamsRecords,writeToSharepoint);
                                Assert.IsNotNull(allShelfmarkFiles);

                            }

                            if (queryMetadata)
                            {
                                bool writeToSharepoint = true;
                                LibraryAPIs.queryMetadataAPIs(spURL, "Digitisation Workflow", SourceFolderStatus,writeToSharepoint);
                            }
                            // ======================================================================

                            // Deprecated?
                            // var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
                            //Assert.IsNotNull(sourceFolderXMLFiles);


                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error while processing query.\n Exception: {0}", ex);
                    }

                    TimeSpan ts = DateTime.Now - startTime;
                    var runTime = ts.TotalSeconds.ToString();
                    Console.WriteLine("=======================================\nFinished in {0} seconds\n=======================================", runTime);
                    Console.WriteLine("Perform another query? (yes/no)");
                    restartProgram = Console.ReadLine();
                }
                else // else you will use an ingest spreadsheet rather than sharepoint
                {
                    // This will get the source folder status automatically (or should do...)

                    try
                    {

                        Console.WriteLine("Searching for an ingest spreadsheet file with name [IngestSpreadsheet.xlsx] and sheet name [PSIP Generator Fields].");
                        var ingestSpreadsheetPath = Directory.GetCurrentDirectory() + "\\IngestSpreadsheet.xlsx";
                        if (!File.Exists(ingestSpreadsheetPath))
                        {
                            Console.WriteLine("Error: Expected to find an ingest spreadsheet in the current directory with name 'IngestSpreadsheet.xlsx'");
                            continue; // this should kick you back to the start of the while loop
                        }

                        //===================== Checks to run ====================

                        bool runImageOrderGenerationXMLChecks = false;
                        bool queryMetadata = false;

                        Console.WriteLine("Generate image order CSV and perform ALTO XML checks? yes/no");
                        String inputImageOrderGen = Console.ReadLine();
                        if (inputImageOrderGen.ToLower().Equals("yes") || inputImageOrderGen.ToLower().Equals("y"))
                        {
                            runImageOrderGenerationXMLChecks = true;
                        }
                        Console.WriteLine("Run query against descriptive metadata APIs? yes/no");
                        String inputMDAPIQuery = Console.ReadLine();
                        if (inputMDAPIQuery.ToLower().Equals("yes") || inputMDAPIQuery.ToLower().Equals("y"))
                        {
                            queryMetadata = true;
                        }
                        // =======================================================


                        // find an alternative to linq-to-excel here...

                        // Need to create a "HMDSPObject" to pass to the relevant bits of the tool and bypass sharepoint...
                        // Turn off writing to sharepoint with a bool or something...
                        List<HMDSPObject> ingestSpreadsheetItems = new List<HMDSPObject>();


                        using (var stream = File.Open(ingestSpreadsheetPath, FileMode.Open, FileAccess.Read))
                        {
                            // Auto-detect format, supports:
                            //  - Binary Excel files (2.0-2003 format; *.xls)
                            //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                            using (var reader = ExcelReaderFactory.CreateReader(stream))
                            {

                                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                                    {
                                        UseHeaderRow = true
                                    }
                                });

                                var table = result.Tables;
                                var resultTable = table["PSIP Generator Fields"];


                                int dummySharepointID = -9999;
                                foreach (System.Data.DataRow dr in resultTable.Rows)
                                {

                                    HMDSPObject item = new HMDSPObject();
                                    item.Title = dr["Shelfmark"].ToString();
                                    item.ID = dummySharepointID.ToString(); // won't have a sharepoint ID
                                    item.Location = dr["Source Folder"].ToString();
                                    item.MetadataSource = dr["Descriptive Metadata Source"].ToString();
                                    item.SystemNumber = dr["System number"].ToString();
                                    dummySharepointID += 1;

                                    ingestSpreadsheetItems.Add(item);

                                }
                            }
                        }


                        var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(ingestSpreadsheetItems);
                        // generates a list of HMDObjects... then the software treats this exactly like it had come from sharepoint
                        if (SourceFolderStatus.Count < 1)
                        {
                            Console.WriteLine("Error: No items found in ingest spreadsheet.\nMake sure the spreadsheet contains the following columns:");
                            Console.WriteLine("[Shelfmark] [Source Folder] [Descriptive Metadata Source] [System number]");
                            Console.WriteLine("Perform another query? (yes/no)");
                            restartProgram = Console.ReadLine();
                        }
                        else
                        {
                            // proceed
                            var env = "prod"; // dictates behaviour in the imageorder csv - find to leave as prod here
                            bool writeToSharepoint = false;
                            var spURL = ""; // make sure we're not using sharepoint at all here
                            if (runImageOrderGenerationXMLChecks)
                            {
                                Console.WriteLine("=======================================\nQuerying metadata APIs to retrieve child shelfmark information...\n=======================================");
                                var iamsRecords = LibraryAPIs.queryMetadataAPIs(spURL, "Digitisation Workflow", SourceFolderStatus, writeToSharepoint);

                                var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus, env, spURL, "Digitisation Workflow", iamsRecords, writeToSharepoint);
                                Assert.IsNotNull(allShelfmarkFiles);

                            }
                            if (queryMetadata)
                            {
                                LibraryAPIs.queryMetadataAPIs(spURL, "Digitisation Workflow", SourceFolderStatus, writeToSharepoint);
                            }

                        }
                    }
                    catch(Exception ex)
                    {
                        Console.WriteLine("Error while processing ingest spreadsheet.\nException: {0}",ex);
                    }
                    TimeSpan ts = DateTime.Now - startTime;
                    var runTime = ts.TotalSeconds.ToString();
                    Console.WriteLine("=======================================\nFinished in {0} seconds\n=======================================", runTime);
                    Console.WriteLine("Perform another query? (yes/no)");
                    restartProgram = Console.ReadLine();
                }
                    

                if (restartProgram.ToLower().Equals("no") || restartProgram.ToLower().Equals("n"))
                {
                    Console.WriteLine("Press any key to exit.");
                    Console.ReadKey();
                    runProgram = false;
                }
                Console.Clear();

            }


                return;
        }

    }

}


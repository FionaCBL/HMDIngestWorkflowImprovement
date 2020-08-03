using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;        
using Microsoft.VisualStudio.TestTools.UnitTesting;


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
            if (env == "prod")
            {
                spURL = "http://hmd.sharepoint.ad.bl.uk";
                project = "Zoroastrian Manuscripts";
            }

            else if (env == "test")
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
            var startTime = DateTime.Now;

            // Set the environment (use 'test' or 'prod')
            var env = "test";
            bool debug = false;

            Console.WriteLine("You are currently running {0}. Switch to prod? yes/no", env);
            String inputEnv = Console.ReadLine();
            if (inputEnv.ToLower() == "yes")
            {
                env = "prod";
            }

            // do initial setup
            Dictionary<String, String> variablesDictionary = InitialSetup(env);
            Assert.IsNotNull(variablesDictionary); // Check we set the variables properly
            var spURL = variablesDictionary["spURL"];
            var project = variablesDictionary["project"];
            String inputVariable = "";
            Assert.IsTrue(RunInitialTests(spURL));

            Console.WriteLine("This program searches sharepoint using either projects or individual shelfmarks. Use shelfmarks? (yes/no)");
            String useShelfmarksYN = Console.ReadLine();
            if (useShelfmarksYN.ToLower() == "yes")
            {
                Console.WriteLine("Enter a shelfmark name (must match Sharepoint site)");
                String inputSingleShelfmark = Console.ReadLine();
                if (inputSingleShelfmark.Length > 0)
                {
                    inputVariable = @"<FieldRef Name ='Title'/><Value Type = 'Text'>" + inputSingleShelfmark + @"</Value>";
                }
                else // kick it back and use defaults if nothing provided
                {
                    Console.WriteLine("No shelfmark provided! Using default project");
                    inputVariable = @"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>";
                }

            }
            else
            {

                Console.WriteLine("Project is currently set to {0}, change this? (yes/no)", project);
                String inputProjectYN = Console.ReadLine();
                if (inputProjectYN.ToLower() == "yes")
                {
                    Console.WriteLine("Type a project name (must match sharepoint record)", project);
                    String inputProject = Console.ReadLine();
                    if (inputProject.Length > 0)
                    {
                        inputVariable = @"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + inputProject + @"</Value>";

                    }
                    else
                    {
                        Console.WriteLine("No project provided! Using default project.");
                        inputVariable = @"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>";

                    }
                }
                else
                {
                    Console.WriteLine("Using default project {0}",project);
                    inputVariable = @"<FieldRef Name ='Project_x0020_Name'/><Value Type = 'Text'>" + project + @"</Value>";

                }
            }

            // Check we assigned something to the input variable
            Assert.AreNotEqual(0, inputVariable.Length);

            //===================== Checks to run ====================
            bool reportShelfmarkCheckStatus = false;
            bool runShelfmarkCharacterChecks = false;
            bool runImageOrderGenerationXMLChecks = false;
            bool queryMetadata = false;


            Console.WriteLine("Run shelfmark source folder checks? yes/no");
            String inputSourceFolderCheck = Console.ReadLine();
            if (inputSourceFolderCheck.ToLower() == "yes")
            {
                reportShelfmarkCheckStatus = true;
            }
            Console.WriteLine("Run shelfmark protected character check? yes/no");
            String inputSMCharCheck = Console.ReadLine();
            if (inputSMCharCheck.ToLower() == "yes")
            {
                runShelfmarkCharacterChecks = true;
            }
            Console.WriteLine("Generate image order CSV and perform ALTO XML checks? yes/no");
            String inputImageOrderGen = Console.ReadLine();
            if (inputImageOrderGen.ToLower() == "yes")
            {
                runImageOrderGenerationXMLChecks = true;
            }
            Console.WriteLine("Run query against descriptive metadata APIs? yes/no");
            String inputMDAPIQuery = Console.ReadLine();
            if (inputMDAPIQuery.ToLower() == "yes")
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
            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListFieldContents(spURL, "Digitisation Workflow",env,inputVariable);
            if(DigitisationWorkflow_ID_Title_SourceFolders.Count < 1)
            {
                Console.ReadKey();
                return;
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
                bool addColumns = false;
                if (addColumns)
                {
                    // Add columns for XML checking
                    String SharePointColumnXMLCheck = "ALTOXMLCheck";
                    Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnXMLCheck));

                }

                var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus, env, spURL, "Digitisation Workflow");
                Assert.IsNotNull(allShelfmarkFiles);

            }

            if (queryMetadata)
            {
                LibraryAPIs.queryMetadataAPIs(spURL,"Digitisation Workflow",SourceFolderStatus);
            }
            // ======================================================================

            // Deprecated?
            // var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
            //Assert.IsNotNull(sourceFolderXMLFiles);

            TimeSpan ts = DateTime.Now - startTime;
            var runTime = ts.TotalSeconds.ToString();
            Console.WriteLine("=======================================\nFinished in {0} seconds\n=======================================",runTime);
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();


            return;            
        }

    }

}


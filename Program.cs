﻿using System;
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
                String testXMLPath = @"G:\Heritage Made Digital\05 Projects\Workflow\Validation of Ingest Workflow\Example OCR files\IOR!P!5644_Jun_1899_nos_80-92_001.xml";
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

            // Set the environment (use 'test' or 'prod')
            var env = "test";
            bool debug = false;


            //===================== Checks to run ====================
            bool reportShelfmarkCheckStatus = false;
            bool runShelfmarkCharacterChecks = false;
            bool runImageOrderGenerationXMLChecks = false;

            Console.WriteLine("Run shelfmark source folder checks? yes/no");
            String inputOne = Console.ReadLine();
            if (inputOne.ToLower() == "yes")
            {
                reportShelfmarkCheckStatus = true;
            }
            Console.WriteLine("Run shelfmark protected character check? yes/no");
            String inputTwo = Console.ReadLine();
            if (inputTwo.ToLower() == "yes")
            {
                runShelfmarkCharacterChecks = true;
            }
            Console.WriteLine("Generate image order CSV and perform ALTO XML checks? yes/no");
            String inputThree = Console.ReadLine();
            if (inputThree.ToLower() == "yes")
            {
                runImageOrderGenerationXMLChecks = true;
            }

            // =======================================================

            // do initial setup
            Dictionary<String, String> variablesDictionary = InitialSetup(env);
            Assert.IsNotNull(variablesDictionary); // Check we set the variables properly
            var spURL = variablesDictionary["spURL"];
            var project = variablesDictionary["project"];

            Assert.IsTrue(RunInitialTests(spURL));

            


            // Add columns for XML checking
            String SharePointColumnXMLCheck = "ALTOXMLCheck";
            Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnXMLCheck));

           

            // Get the 'Digitisation Workflow' list fields and print them out:
            if (debug)
            {
                var DigitisationWorkflowTitles = SharepointTools.GetSharePointListFields(spURL, "Digitisation Workflow");
                Assert.IsNotNull(DigitisationWorkflowTitles.Count);
            }

            // ============== Preliminary stuff - has to be run every time ================
            // This is the "grab all the info from sharepoint" part of things

            // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list
            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListFieldContents(spURL, "Digitisation Workflow",env,project);
            Assert.IsNotNull(DigitisationWorkflow_ID_Title_SourceFolders.Count);

            //  Check source folders - requires the above two lines to work
            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
            Assert.IsNotNull(SourceFolderStatus.Count);
            if (!debug)
            {
                Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus, "sourceFolderStatus"));
            }

            // =============================================================================

            // ================= Optional checks - some take ~10 mins to read all of sharepoint, so don't run every time =============
            
            // Validates that the value of the 'Source Folder' field exists as a valid path on the network
            // Attempts to write the results to a sharepoint column for each row (== shelfmark)
            if (reportShelfmarkCheckStatus)
            {
                String SharePointSourceFolderCheck = "SourceFolderValid";
                Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointSourceFolderCheck));
                Assert.IsTrue(SharepointTools.ReportSourceFolderStatus(spURL, "Digitisation Workflow", SharePointSourceFolderCheck, SourceFolderStatus));
            }

            // Checks for protected characters in Shelfmark names 
            // Attempts to write the results to sharepoint column
            if (runShelfmarkCharacterChecks)
            {
                // Add in Shelfmark protected chars check here
                String SharePointColumnShelfmarkCheck = "ShelfmarkCheck";
                Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck));
                List<List<String>> badShelfmarks = SharepointTools.BadShelfmarkNames(SourceFolderStatus);
                if (badShelfmarks.Count != 0)
                {
                    // now we need to write to sharepoint by shelfmark
                    String badShelfmarkMessage = "Protected character found in shelfmark";
                    //Assert.IsTrue(SharepointTools.WriteToSharepointColumnByShelfmark(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck, badShelfmarks));
                    // Try by ID instead...
                    foreach (List<String> item in badShelfmarks)
                    {
                        Int32 ID = Int32.Parse(item[0]);
                        String SM = item[1];
                        Assert.IsTrue(SharepointTools.WriteToSharepointColumnByID(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck, SM, ID, badShelfmarkMessage));
                    }
                }
            }

            // Get the labels (image order, image type etc) for all shelfmarks passed into this function
            // Write these out in an image order csv 
            // Also runs ALTO XML checks - do they exist with the same names as TIFs and in the same number as TIFs?
            // Are they version 2.0 or older?
            if (runImageOrderGenerationXMLChecks)
            {
                var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus, env, spURL, "Digitisation Workflow");
                Assert.IsNotNull(allShelfmarkFiles);
            }
            // ======================================================================

            // Deprecated?
            // var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
            //Assert.IsNotNull(sourceFolderXMLFiles);




            return;            
        }

    }

}


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
        static void Main(string[] args)
        {

            // Set the environment
            var spURL = "";
            var env = "test";
            var project = "";

            Assert.IsNotNull(env);

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
                return;
            }

            // test a single XML version number
            bool goodXMLVersion = false;
            String testXMLPath = @"G:\Heritage Made Digital\05 Projects\Workflow\Validation of Ingest Workflow\Example OCR files\IOR!P!5644_Jun_1899_nos_80-92_001.xml";
            String XMLVersionNumber = InputOrderSpreadsheetTools.GetXMLVersionNumber(testXMLPath);
            float XMLVNum = float.Parse(XMLVersionNumber,System.Globalization.CultureInfo.InvariantCulture);
            if (XMLVNum <= 2.0)  
            {
                goodXMLVersion = true;
            }
            Assert.IsTrue(goodXMLVersion);

            // Add columns for XML checking
            String SharePointColumnXMLCheck = "ALTOXMLCheck";
            Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointColumnXMLCheck));


            // Test sharepoint site exists
            Assert.IsTrue(SharepointTools.SharepointSiteExists(spURL));

            // Tests that you can retrieve a tile from the sharepoint site
            var spTitle = SharepointTools.GetSharepointTitle(spURL);
            Assert.IsNotNull(spTitle);
            Console.WriteLine("Sharepoint site title: {0}",spTitle);
            var spListNames = SharepointTools.GetSharePointLists(spURL);
            Assert.IsTrue(spListNames.Count != 0);

            // Get the 'Digitisation Workflow' list fields and print them out:
            // ONLY FOR DEBUGGING PURPOSES:
            
            var DigitisationWorkflowTitles = SharepointTools.GetSharePointListTitles(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflowTitles.Count);
            


            // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list
            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListTitleContents(spURL, "Digitisation Workflow",env,project);
            Assert.IsNotNull(DigitisationWorkflow_ID_Title_SourceFolders.Count);

            //  Check source folders - requires the above two lines to work
            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
            Assert.IsNotNull(SourceFolderStatus.Count);
            Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus,"sourceFolderStatus"));

            String SharePointSourceFolderCheck = "SourceFolderValid";
            Assert.IsTrue(SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", SharePointSourceFolderCheck));

            Assert.IsTrue(SharepointTools.ReportSourceFolderStatus(spURL, "Digitisation Workflow", SharePointSourceFolderCheck, SourceFolderStatus));

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
                    Assert.IsTrue(SharepointTools.WriteToSharepointColumnByID(spURL, "Digitisation Workflow", SharePointColumnShelfmarkCheck,SM, ID,badShelfmarkMessage));
                }
                }
            // can write this out at the end in a little console report


            var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
           // this should never be null, even if every list within sourceFolderXMLFiles is null...
            Assert.IsNotNull(sourceFolderXMLFiles);
       
            
            // Park this for now. You do not need to write out which XMLs were found.
            //Assert.IsTrue(TextOutputFunctions.OutputListOfLists(sourceFolderXMLFiles, "xmlFilesFound"));

            // Pass in some information wrt shelfmark and source folder status, then searches for the TIF files in each of the folders


            // Get the labels (image order, image type etc) for all shelfmarks passed into this function
            var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus,env, spURL,"Digitisation Workflow");
            Assert.IsNotNull(allShelfmarkFiles);
            

            // Compare number of XML files (if any) to number of tifs...
            // Maybe put the XML getting functionality here?






            //Assert.IsTrue(InputOrderSpreadsheetTools.RetrieveImgOrderLabels(allShelfmarkFiles,env));

            return;            
        }

    }

}


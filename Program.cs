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
                project = "Zoroastrian Manuscripts"; // exists in the test SP too - for now
            }
            else
            {
                Console.WriteLine("You must set an environment variable.");
                return;
            }
            
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
            /*
            var DigitisationWorkflowTitles = SharepointTools.GetSharePointListTitles(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflowTitles.Count);
            */


            // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list
            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListTitleContents(spURL, "Digitisation Workflow",env,project);
            Assert.IsNotNull(DigitisationWorkflow_ID_Title_SourceFolders.Count);

            //  Check source folders - requires the above two lines to work
            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
            Assert.IsNotNull(SourceFolderStatus.Count);
            Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus,"sourceFolderStatus"));



            var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
           // this should never be null, even if every list within sourceFolderXMLFiles is null...
            Assert.IsNotNull(sourceFolderXMLFiles);
       
            
            // Park this for now. You do not need to write out which XMLs were found.
            //Assert.IsTrue(TextOutputFunctions.OutputListOfLists(sourceFolderXMLFiles, "xmlFilesFound"));

            // Pass in some information wrt shelfmark and source folder status, then searches for the TIF files in each of the folders


            // Get the labels (image order, image type etc) for all shelfmarks passed into this function
            var allShelfmarkFiles = InputOrderSpreadsheetTools.listAllShelfmarkFilesTIFXML(SourceFolderStatus,env);
            Assert.IsNotNull(allShelfmarkFiles);

            // Compare number of XML files (if any) to number of tifs...
            // Maybe put the XML getting functionality here?






            //Assert.IsTrue(InputOrderSpreadsheetTools.RetrieveImgOrderLabels(allShelfmarkFiles,env));

            return;            
        }

    }

}


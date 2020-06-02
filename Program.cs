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

            Assert.IsNotNull(env);

            if (env == "prod")
            {
                spURL = "http://hmd.sharepoint.ad.bl.uk";
            }

            else if (env == "test")
            {
                spURL = "http://v12t-sp13wfe1:88/";

            }
            else
            {
                Console.WriteLine("You must set an environment variable.");
                return;
            }

            // Start off just testing if things pass imaging and conservation stages

            // Imaging Status = Complete



            // Tests that you can retrieve a tile from the sharepoint site
            Assert.IsTrue(SharepointTools.SharepointSiteExists(spURL));

            var spTitle= SharepointTools.GetSharepointTitle(spURL);
            Assert.IsNotNull(spTitle);
            Console.WriteLine("Sharepoint site title: {0}",spTitle);
            var spListNames = SharepointTools.GetSharePointLists(spURL);
            Assert.IsTrue(spListNames.Count != 0);

            // Get the 'Digitisation Workflow' list contents:
            var DigitisationWorkflowTitles = SharepointTools.GetSharePointListTitles(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflowTitles.Count);

            // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list
            var DigitisationWorkflow_ID_Title_SourceFolders = SharepointTools.GetSharePointListTitleContents(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflow_ID_Title_SourceFolders.Count);

            // At this point you could add in some filtering by shelfmark, otherwise it will return all shelfmarks

            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
            Assert.IsNotNull(SourceFolderStatus.Count);
            Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus,"sourceFolderStatus"));

            //Getting & writing out the XML is currently broken, swing back around and fix this (takes hours)
            //var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
            //Assert.IsNotNull(sourceFolderXMLFiles);
            // Assert.IsTrue(TextOutputFunctions.OutputListOfLists(sourceFolderXMLFiles, "xmlFilesFound"));

            // Pass in some information wrt shelfmark and source folder status, then searches for the TIF files in each of the folders


            // Get the image order spreadsheet
            var allShelfmarkFiles = InputOrderSpreadsheetTools.getAllShelfmarkTIFs(SourceFolderStatus,env);
            Assert.IsNotNull(allShelfmarkFiles);

            Assert.IsTrue(InputOrderSpreadsheetTools.RetrieveImgOrderLabels(allShelfmarkFiles,env));

            return;            
        }

    }

}


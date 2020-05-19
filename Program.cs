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
            var environment = "test";
            var spURL = "";

            if (environment == "prod")
            {
                spURL = "http://hmd.sharepoint.ad.bl.uk";
            }

            else if (environment == "test")
            {
                spURL = "http://v12t-sp13wfe1:88/";

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

            var SourceFolderStatus = SharepointTools.CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);
            Assert.IsNotNull(SourceFolderStatus.Count);

            Assert.IsTrue(TextOutputFunctions.OutputListOfLists(SourceFolderStatus,"sourceFolderStatus"));

            //Getting & writing out the XML is currently broken, swing back around and fix this (takes hours)
            //var sourceFolderXMLFiles = SharepointTools.GetSourceFolderXMLs(SourceFolderStatus, true);
            //Assert.IsNotNull(sourceFolderXMLFiles);
            // Assert.IsTrue(TextOutputFunctions.OutputListOfLists(sourceFolderXMLFiles, "xmlFilesFound"));


            // Get the image order spreadsheet
            var allShelfmarkFiles = InputOrderSpreadsheetTools.getAllShelfmarkTIFs(SourceFolderStatus);
            Assert.IsNotNull(allShelfmarkFiles);

            Assert.IsTrue(InputOrderSpreadsheetTools.RetrieveImgOrderLabels(allShelfmarkFiles));



            return;
        }

    }

}


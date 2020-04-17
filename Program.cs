using System;
using System.Web;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;        
using System.Net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint;


namespace HMDSharepointChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            var spURL = "http://hmd.sharepoint.ad.bl.uk";

            // Tests that you can retrieve a tile from the sharepoint site
            Assert.IsTrue(SharepointSiteExists(spURL));

            var spTitle= GetSharepointTitle(spURL);
            Assert.IsNotNull(spTitle);
            Console.WriteLine("Sharepoint site title: {0}",spTitle);
            var spListNames = GetSharePointLists(spURL);
            Assert.IsTrue(spListNames.Count != 0);

            // Get the 'Digitisation Workflow' list contents:
            var DigitisationWorkflowTitles = GetSharePointListTitles(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflowTitles.Count);

            // Get the contents of the "Title" column in the 'Digitisation Workflow' list
            var DigitisationWorkflowTitleContent = GetSharePointListTitleContents(spURL, "Digitisation Workflow", "Shelfmark","Title");
            Assert.IsNotNull(DigitisationWorkflowTitleContent.Count);

            return;
        }

        private static bool SharepointSiteExists(string url)
        {
            try
            {

                ClientContext clientContext = new ClientContext(url);
                Web site = clientContext.Web;
                clientContext.Load(site);
                clientContext.ExecuteQuery();
                var siteTitle = site.Title;
                return !string.IsNullOrEmpty(siteTitle);
       
            }
            catch
            {
                // Any exception returns false
                return false;
            }

        }

        private static string GetSharepointTitle(string sharepointSite)
        {

            try
            {
                ClientContext clientContext = new ClientContext(sharepointSite);
                Web site = clientContext.Web;
                clientContext.Load(site);
                clientContext.ExecuteQuery();
                var siteTitle = site.Title;
                return siteTitle;
            }
            catch
            {
                return null;
            }
        }

        private static List<String> GetSharePointLists(string sURL)
        {

            try
            {
                ClientContext clientContext = new ClientContext(sURL);
                Web oSite = clientContext.Web;
                ListCollection collList = oSite.Lists;

                clientContext.Load(collList);
                clientContext.ExecuteQuery();


                List<string> listNames = new List<string>();
                foreach (SP.List oList in collList)
                {
                    Console.WriteLine("Title: {0} Created: {1}", oList.Title, oList.Created.ToString());
                    listNames.Add(oList.Title);
                }
                return listNames;
            }
            catch
            {
                return null;
            }
        }

        private static List<String> GetSharePointListTitles(string sURL, string lName)
        {

            try
            {

                
                ClientContext clientContext = new ClientContext(sURL);
                SP.List oList = clientContext.Web.Lists.GetByTitle(lName);

                if (oList != null)
                {

                    clientContext.Load(oList.Fields);
                    clientContext.ExecuteQuery();

                    List<string> listColumns = new List<string>();

                    foreach (SP.Field myField in oList.Fields)
                    {

                        var thingToPrint = myField.Title+", " +myField.InternalName;
                        //Console.WriteLine(myField.Title);
                        Console.WriteLine(thingToPrint);

                        listColumns.Add(myField.Title);

                    }

                    return listColumns;
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        private static List<String> GetSharePointListTitleContents(string sURL, string lName, string tName, string iTName)
        {
            var myID = "";
            var myTitle = "";
            var myLoc = "";

            try
            {

                ClientContext clientContext = new ClientContext(sURL);
                SP.List oList = clientContext.Web.Lists.GetByTitle(lName);
                //CamlQuery camlQuery = new CamlQuery();
                //camlQuery.ViewXml = "<Where><IsNotNull><FieldRef Name='Source Folder'/></IsNotNull></Where>";
                SP.ListItemCollection oItems = oList.GetItems(CamlQuery.CreateAllItemsQuery());

                clientContext.Load(oList);
                clientContext.Load(oItems);
                clientContext.ExecuteQuery();

                List<string> listRows = new List<string>();


                foreach (Microsoft.SharePoint.Client.ListItem oListItem in oItems)
                {

                    var itemID = oListItem.FieldValues["ID"].ToString();
                    var itemTitle = oListItem.FieldValues["Title"].ToString();
                    var itemLocation = "";
                    try
                    {
                        itemLocation = ((Microsoft.SharePoint.Client.FieldUrlValue)(oListItem["Source_x0020_Folder0"])).Url.ToString();
                    }

                    catch
                    {
                        continue; // If the itemLocation is empty, we don't care, but this throws an exception so need to skip over this item
                    }
                    if (itemLocation != null)
                    {
                        String rowString = String.Format("ID: {0} \t Title: {1} \t Location: {2}", itemID, itemTitle, itemLocation);
                        myID = itemID;
                        myTitle = itemTitle;
                        myLoc = itemLocation;
                        Console.WriteLine(rowString);
                        listRows.Add(rowString);

                    }
                    
                }

                return listRows;

            }
            catch
            {
                var theID = myID;
                var theTitle = myTitle;
                var theLoc = myLoc;
                return null;
            }
        }


    }
}

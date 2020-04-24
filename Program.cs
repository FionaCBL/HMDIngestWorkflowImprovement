using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;        
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;


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

            // Get the contents of the "ID", "Shelfmark" and "Source Folder" columns in the 'Digitisation Workflow' list
            var DigitisationWorkflow_ID_Title_SourceFolders = GetSharePointListTitleContents(spURL, "Digitisation Workflow");
            Assert.IsNotNull(DigitisationWorkflow_ID_Title_SourceFolders.Count);

            var SourceFolderStatus = CheckSourceFolderExists(DigitisationWorkflow_ID_Title_SourceFolders);


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
                    //Console.WriteLine("Title: {0} Created: {1}", oList.Title, oList.Created.ToString());
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
                        //Console.WriteLine(thingToPrint);

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

        private static List<List<String>> GetSharePointListTitleContents(string sURL, string lName)
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


                
                List<List<string>> listAll = new List<List<string>>();

                foreach (Microsoft.SharePoint.Client.ListItem oListItem in oItems)
                {
                    List<string> listItem = new List<string>();

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
                        //Console.WriteLine(rowString);
                        listItem.Add(myID);
                        listItem.Add(myTitle);
                        listItem.Add(myLoc);

                    }
                    listAll.Add(listItem);
                }

                return listAll;

            }
            catch
            {
                var theID = myID;
                var theTitle = myTitle;
                var theLoc = myLoc;
                return null;
            }
        }


        private static List<List<String>> CheckSourceFolderExists(List<List<string>> itemList)
        {

            // Need to translate the source folder paths retrieved from Sharepoint
            // into the actual source folder locations including the shelfmarks
            // Shelfmarks need to be transformed as per the DIPS naming requirements



            List<List<String>> folderExistenceStatus = new List<List<String>>();


            foreach (var item in itemList)
            {
                List<String> itemStatus = new List<String>();

                //Console.WriteLine("{0} \t {1} \t {2}", item[0], item[1],item[2]);
                string ID = item[0];
                string Shelfmark = item[1];
                string sourceFolderSP = item[2];
                string sourceFolder = sourceFolderSP.Replace(@"////", @"//");
                sourceFolder = sourceFolderSP.Replace(@"/", @"\");
                sourceFolder = sourceFolder.Replace(@"file:", @"");

                var sfAlt2 = sourceFolder.Split('\\')[2];
                try
                {
                    string sfAlt = sourceFolder.Replace(sfAlt2, @"ad\collections");
                    bool DirectoryExists = false;
                    bool altDirectoryExists = false;
                    if (Directory.Exists(sourceFolder))
                    {
                        DirectoryExists = true;
                    }
                    else
                    {
                        altDirectoryExists = Directory.Exists(sfAlt);

                        //if (DirectoryExists)
                        //{

                        //string realFolderLocation = UNCPath(sfAlt);
                        // string share = Dfs.GetDfsInfo(sfAlt);

                        // Need to put in DFS finding stuff here, but not now


                        //}

                    }

                    if (DirectoryExists) Console.WriteLine("Folder: {0} \t Exists: {1}", sourceFolder, DirectoryExists);
                    else if (altDirectoryExists)
                    {
                        Console.WriteLine("Folder: {0} \t Exists at {1}: {2}", sourceFolder, sfAlt, altDirectoryExists);
                    }
                    else
                    {
                        Console.WriteLine("ERROR: Folder {0} not found", sourceFolder);
                    }
                    string folderStatus = DirectoryExists.ToString();
                    itemStatus.Add(ID);
                    itemStatus.Add(Shelfmark);
                    itemStatus.Add(sourceFolder);
                    itemStatus.Add(folderStatus);
                    itemStatus.Add(altDirectoryExists.ToString());

                    folderExistenceStatus.Add(itemStatus);
                    }
                catch
                {
                    return null;
                    // really need to handle this exception properly!
                }
                }

            // blah

            return folderExistenceStatus;

        }
        
        /*
         public static string UNCPath(string path)
        {
            if (!path.StartsWith(@"\\"))
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Network\\" + path[0]))
                {
                    if (key != null)
                    {
                        return key.GetValue("RemotePath").ToString() + path.Remove(0, 2).ToString();
                    }
                }
            }
            return path;
        }
        */

    }



}

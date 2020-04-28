using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    public class SharepointTools
    {
        public static bool SharepointSiteExists(string url)
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

        public static string GetSharepointTitle(string sharepointSite)
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

        public static List<String> GetSharePointLists(string sURL)
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


        public static List<String> GetSharePointListTitles(string sURL, string lName)
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

                        var thingToPrint = myField.Title + ", " + myField.InternalName;
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



        public static List<List<String>> GetSharePointListTitleContents(string sURL, string lName)
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


        public static List<List<String>> CheckSourceFolderExists(List<List<string>> itemList)
        {

            // Need to translate the source folder paths retrieved from Sharepoint
            // into the actual source folder locations including the shelfmarks
            // Shelfmarks need to be transformed as per the DIPS naming requirements


            List<List<String>> folderExistenceStatus = new List<List<String>>();
            List<String> fESTitles = new List<String>();
            fESTitles.Add("ID");
            fESTitles.Add("Shelfmark");
            fESTitles.Add("Source Folder");
            fESTitles.Add("Directory Status");
            fESTitles.Add("Alt-Directory Status");
            fESTitles.Add("Source Folder Error");
            folderExistenceStatus.Add(fESTitles); // Add the titles list as the first item in the list of lists

            foreach (var item in itemList)
            {
                List<String> itemStatus = new List<String>();
                bool checkSourceFolder = false;

                string fullSourceFolderPath = "";

                //Console.WriteLine("{0} \t {1} \t {2}", item[0], item[1],item[2]);
                string ID = item[0];
                string Shelfmark = item[1];
                string sourceFolderSP = item[2];
                string sourceFolder = sourceFolderSP.Replace("////", "//");
                var sf1 = sourceFolder;
                sourceFolder = sourceFolder.Replace("/", @"\");
                var sf2 = sourceFolder;
                sourceFolder = sourceFolder.Replace(@"file:", @"");
                var sf3 = sourceFolder;
                if (sourceFolder.Contains(@"\\\"))
                {
                    sourceFolder = sourceFolder.Replace(@"\\\", @"\\"); // this is there in some cases...
                    var sf4 = sourceFolder;
                    // Can flag this 
                    checkSourceFolder = true;
                }


                var sfAlt2 = sourceFolder.Split('\\')[2]; // Get the part of the string with server name in

                try
                {
                    string sfAlt = sourceFolder.Replace(sfAlt2, @"ad\collections");
                    bool DirectoryExists = false;
                    bool altDirectoryExists = false;
                    if (Directory.Exists(sourceFolder))
                    {
                        DirectoryExists = true;
                        fullSourceFolderPath = ConstructFullFolderName(Shelfmark, sourceFolder);
                    }
                    else
                    {
                        altDirectoryExists = Directory.Exists(sfAlt);
                        fullSourceFolderPath = ConstructFullFolderName(Shelfmark, sfAlt);

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
                        checkSourceFolder = true;
                    }
                    else
                    {
                        Console.WriteLine("ERROR: Folder {0} not found", sourceFolder);
                        checkSourceFolder = true;
                    }
                    string folderStatus = DirectoryExists.ToString();
                    itemStatus.Add(ID);
                    itemStatus.Add(Shelfmark);
                    itemStatus.Add(sourceFolder);
                    itemStatus.Add(folderStatus);
                    itemStatus.Add(altDirectoryExists.ToString());
                    itemStatus.Add(checkSourceFolder.ToString());
                    itemStatus.Add(fullSourceFolderPath);

                }
                catch
                {
                    bool DirectoryStatus = false;
                    bool altDirectoryStatus = false;
                    checkSourceFolder = true;
                    itemStatus.Add(ID);
                    itemStatus.Add(Shelfmark);
                    itemStatus.Add(sourceFolder);
                    itemStatus.Add(DirectoryStatus.ToString());
                    itemStatus.Add(altDirectoryStatus.ToString());
                    itemStatus.Add(checkSourceFolder.ToString());
                    itemStatus.Add(fullSourceFolderPath);


                    // really need to handle this exception properly!
                }
                folderExistenceStatus.Add(itemStatus);


                // Need to decide how to do reporting with this - final bool shows whether the output folder needs flagging.
                // Source folder needs flagging if: source folder not found at all, or source folder found under \\ad\collections but not under the location given


            }
            return folderExistenceStatus;

        }


        private static String ConstructFullFolderName(string SM, string sF)
        {
            string fullPath = null;
            string SM_folderFormat = SM.ToLower().Replace(@" ", @"_").Replace(@"/",@"!").Replace(@".",@"_");

            if( sF.Contains(SM_folderFormat))
            {
                // do nothing! this is the ideal scenario
                fullPath = sF;
            }

            else
            {
                var testPath = "";
                if (sF.EndsWith(@"\"))
                {
                    testPath = sF + SM_folderFormat;

                }
                else
                {
                    testPath = sF + @"\" + SM_folderFormat;
                }
                
                if (Directory.Exists(testPath))
                {
                    fullPath = testPath;
                }
                
            }

            return fullPath;
        }


    }




}

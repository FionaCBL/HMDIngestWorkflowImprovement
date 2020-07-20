using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    public class HMDSPObject // As retrieved from the HMD sharepoint site
    {
        public string ID { get; set; }
        public string Title { get; set; } 
        public string Location { get; set; }

        public HMDSPObject(string id, string title, string loc)
        {
            ID = id;
            Title = title;
            Location = loc;
        }
    }

    public class HMDObject // adding additional information to the things retrieved from the HMD sharepoint site
    {
        public string ID { get; set; }
        public string Shelfmark { get; set; }
        public string SourceFolderPath { get; set; }
        public string FolderStatus { get; set; }
        public bool AltDirectoryExists { get; set; }
        public bool SourceFolderValid { get; set; }
        public string FullSourceFolderPath { get; set; }
        public bool SourceFolderPathValidElsewhere { get; set; }

        public bool BadShelfmark { get; set; }


        public HMDObject(string id, string shelfmark, string sfpath, string folderstatus, bool adexists, bool sfvalid, string fsfpath, bool altvalid, bool badSM)
        {
            ID = id;
            Shelfmark = shelfmark;
            SourceFolderPath = sfpath;
            FolderStatus = folderstatus;
            AltDirectoryExists = adexists;
            SourceFolderValid = sfvalid;
            FullSourceFolderPath = fsfpath;
            SourceFolderPathValidElsewhere = altvalid;
            BadShelfmark = badSM;

        }
        public HMDObject()
        {
            ID = null;
            Shelfmark = null;
            SourceFolderPath = null;
            FolderStatus = null;
            AltDirectoryExists = false;
            SourceFolderValid = false;
            FullSourceFolderPath = null;
            SourceFolderPathValidElsewhere = false;
            BadShelfmark = false;
        }

    }
    public class SharepointTools
    {
        public static bool SharepointSiteExists(string url)
        {
            using (ClientContext ctx = new ClientContext(url))
            {
                try
                {

                    Web site = ctx.Web;
                    ctx.Load(site);
                    ctx.ExecuteQuery();
                    var siteTitle = site.Title;
                    return !string.IsNullOrEmpty(siteTitle);

                }

                catch (Exception ex)
                {
                    Console.WriteLine("Failed to find sharepoint site. " + ex.Message);

                    // Any exception returns false

                    return false;
                }
            }

        }

        public static string GetSharepointTitle(string sharepointSite)
        {

            try
            {
                using (ClientContext ctx = new ClientContext(sharepointSite))
                {
                    Web site = ctx.Web;
                    ctx.Load(site);
                    ctx.ExecuteQuery();
                    var siteTitle = site.Title.TrimEnd();
                    return siteTitle;
                }
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
                using (ClientContext clientContext = new ClientContext(sURL))
                {
                    Web oSite = clientContext.Web;
                    ListCollection collList = oSite.Lists;

                    clientContext.Load(collList);
                    clientContext.ExecuteQuery();


                    List<string> listNames = new List<string>();
                    foreach (SP.List oList in collList)
                    {
                        listNames.Add(oList.Title);
                    }
                    return listNames;
                }
            }
            catch
            {
                return null;
            }
        }


        public static List<String> GetSharePointListFields(string sURL, string lName)
        {

            try
            {


                using (ClientContext clientContext = new ClientContext(sURL))
                {
                    SP.List oList = clientContext.Web.Lists.GetByTitle(lName);

                    if (oList != null)
                    {

                        clientContext.Load(oList.Fields);
                        clientContext.ExecuteQuery();

                        List<string> listColumns = new List<string>();

                        foreach (SP.Field myField in oList.Fields)
                        {

                            var thingToPrint = myField.Title + ", " + myField.InternalName;
                            Console.WriteLine(thingToPrint); // print this to get the internal name of columns

                            listColumns.Add(myField.Title);

                        }

                        return listColumns;
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            catch
            {
                return null;
            }
        }



        public static List<HMDSPObject> GetSharePointListFieldContents(string sURL, string lName, string env, string inputvar)
        {
            Console.Clear();
            Console.WriteLine("=======================================\nRetrieving sharepoint list items\n=======================================");
            try
            {
                using (ClientContext clientContext = new ClientContext(sURL))
                {

                    SP.List oList = clientContext.Web.Lists.GetByTitle(lName);
                    CamlQuery camlQuery = new CamlQuery();
                    var myQuery = @"<View><Query><Where><Contains>" + inputvar + @"</Contains></Where></Query></View>";
                    camlQuery.ViewXml = String.Format(myQuery);
                    ListItemCollection oItems = oList.GetItems(camlQuery);
                    clientContext.Load(oItems);
                    clientContext.ExecuteQuery();


                    List<HMDSPObject> itemsFound = new List<HMDSPObject>();

                    var itemCounter = 1;
                    foreach (Microsoft.SharePoint.Client.ListItem oListItem in oItems)
                    {
                        if (oItems.Count > 20)
                        {
                            if (itemCounter % 10 == 0)
                            {
                                Console.WriteLine("{0}/{1}", itemCounter, oItems.Count);
                            }
                        }
                        else
                        {
                            Console.WriteLine("{0}/{1}", itemCounter, oItems.Count);

                        }
                        itemCounter += 1;

                        List<string> listItem = new List<string>();
                        HMDSPObject thisItem = new HMDSPObject("", "", ""); // create empty hmdobject


                        var itemID = oListItem.FieldValues["ID"].ToString();
                        var itemTitle = oListItem.FieldValues["Title"].ToString();
                        var itemLocation = "";
                        try
                        {
                            itemLocation = ((Microsoft.SharePoint.Client.FieldUrlValue)(oListItem["Source_x0020_Folder0"])).Url.ToString();
                        }

                        catch (Exception ex)
                        {
                            Console.WriteLine("Could not retrieve 'Source Folder' entry for shelfmark {0}\nException: {1}", itemTitle, ex);
                            thisItem.ID = itemID;
                            thisItem.Title = itemTitle;
                            thisItem.Location = null; // Still want to write out null values of item location so we can report in sharepoint later!
                            itemsFound.Add(thisItem);

                            continue; // If the itemLocation is empty, we don't care, but this throws an exception so need to skip over this item


                        }
                        if (itemLocation != null)
                        {
                            String rowString = String.Format("ID: {0} \t Project: {1} \t Title: {2} \t Location: {3}", itemID, oListItem.FieldValues["Project_x0020_Name"].ToString(), itemTitle, itemLocation);
                            //Console.WriteLine(rowString);
                            thisItem.ID = itemID;
                            thisItem.Title = itemTitle;
                            thisItem.Location = itemLocation;

                        }

                        itemsFound.Add(thisItem);
                    }

                    return itemsFound;

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error - exception {0}", ex);
                return null;
            }

        }



        public static List<HMDObject> CheckSourceFolderExists(List<HMDSPObject> itemList)
        {

            // Need to translate the source folder paths retrieved from Sharepoint
            // into the actual source folder locations including the shelfmarks
            // Shelfmarks need to be transformed as per the DIPS naming requirements


            List<HMDObject> folderExistenceStatus = new List<HMDObject>();
            Console.WriteLine("=======================================\nValidating Source Folder Paths\n=======================================");
            var thisItem = 1;

            foreach (var item in itemList)
            {
                if (itemList.Count > 20)
                {


                    if (thisItem % 10 == 0)
                    {
                        Console.WriteLine("Processing item {0} of {1}", thisItem, itemList.Count);
                    }
                }
                else
                {
                    Console.WriteLine("Processing item {0} of {1}", thisItem, itemList.Count);

                }
                thisItem += 1;
                HMDObject HMDItem = new HMDObject(); // initialise new HMD object with null vals
                    
                bool sourceFolderValid = false
                    ;
                bool sourceFolderValidElsewhere = false;

                string fullSourceFolderPath = "";

                string ID = item.ID;
                string Shelfmark = item.Title;
                string sourceFolderSP = item.Location;

                if (!String.IsNullOrEmpty(sourceFolderSP))
                {

                    string sourceFolder = sourceFolderSP.Replace("////", "//");
                    sourceFolder = sourceFolder.TrimEnd(); // trims whitespace from end
                    var sf1 = sourceFolder;
                    sourceFolder = sourceFolder.Replace("/", @"\");
                    var sf2 = sourceFolder;
                    sourceFolder = sourceFolder.Replace(@"file:", @"");
                    var sf3 = sourceFolder;
                    if (sourceFolder.Contains(@"\\\"))
                    {
                        // I haven't seen a case where this makes the source folder fail yet, so this isn't fatal
                        sourceFolder = sourceFolder.Replace(@"\\\", @"\\"); // this is there in some cases...
                        var sf4 = sourceFolder;
                    }
                    if (sourceFolder.Contains(@"%20"))
                    {
                        sourceFolder = sourceFolder.Replace(@"%20", @" ");
                    }

                    var sfAlt2 = sourceFolder.Split('\\')[2]; // Get the part of the string with server name in

                    try
                    {
                        string sfAlt = sourceFolder.Replace(sfAlt2, @"ad\collections");
                        bool DirectoryExists = false;
                        bool altDirectoryExists = false;
                        if (Directory.Exists(sourceFolder)) // Optimal case - the URL in sharepoint is correct!
                        {
                            sourceFolderValid = true;
                            DirectoryExists = true;
                            fullSourceFolderPath = ConstructFullFolderName(Shelfmark, sourceFolder);
                        }
                        else
                        {
                            altDirectoryExists = Directory.Exists(sfAlt);
                            if (altDirectoryExists) // Next most-optimal case - URL in sharepoint is wrong but it's just the server
                            {
                                sourceFolderValidElsewhere = true;
                                fullSourceFolderPath = ConstructFullFolderName(Shelfmark, sfAlt);

                            }
                            else
                            {
                                fullSourceFolderPath = "";
                            }


                        }

                        if (!DirectoryExists && altDirectoryExists) 
                        {
                            Console.WriteLine("Folder: {0} \t Exists at {1}: {2}", sourceFolder, sfAlt, altDirectoryExists);
                        }
                        else if (!DirectoryExists && !altDirectoryExists)
                        {
                            Console.WriteLine("ERROR: Folder {0} not found", sourceFolder);

                        }
                        string folderStatus = DirectoryExists.ToString();
                        HMDItem.ID = ID;
                        HMDItem.Shelfmark = Shelfmark;
                        HMDItem.SourceFolderPath = sourceFolder;
                        HMDItem.FolderStatus = folderStatus;
                        HMDItem.AltDirectoryExists = altDirectoryExists;
                        HMDItem.SourceFolderValid = sourceFolderValid;
                        HMDItem.FullSourceFolderPath = fullSourceFolderPath;
                        HMDItem.SourceFolderPathValidElsewhere = sourceFolderValidElsewhere;


                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR: Shelfmark {0}\n Exception: {1}", Shelfmark, ex);
                        bool DirectoryStatus = false;
                        bool altDirectoryStatus = false;
                        HMDItem.ID = ID;
                        HMDItem.Shelfmark = Shelfmark;
                        HMDItem.SourceFolderPath = sourceFolder;
                        HMDItem.FolderStatus = DirectoryStatus.ToString();
                        HMDItem.AltDirectoryExists = altDirectoryStatus;
                        HMDItem.SourceFolderValid = sourceFolderValid;
                        HMDItem.FullSourceFolderPath = fullSourceFolderPath;
                        HMDItem.SourceFolderPathValidElsewhere = sourceFolderValidElsewhere;
                        folderExistenceStatus.Add(HMDItem);
                        continue;

                        // really need to handle this exception properly!
                    }
                }
                else // in case you're passed something with a null value for source folder
                {
                    bool DirectoryStatus = false;
                    bool altDirectoryStatus = false;
                    HMDItem.ID = ID;
                    HMDItem.Shelfmark = Shelfmark;
                    HMDItem.SourceFolderPath = null;
                    HMDItem.FolderStatus = DirectoryStatus.ToString();
                    HMDItem.AltDirectoryExists = altDirectoryStatus;
                    HMDItem.SourceFolderValid = false;
                    HMDItem.FullSourceFolderPath = null;
                    HMDItem.SourceFolderPathValidElsewhere = false;
                }
                folderExistenceStatus.Add(HMDItem);

                
            }
            return folderExistenceStatus;

        }

        public static bool ReportSourceFolderStatus(string spURL, string spList, string SFCol, List<HMDObject> SFStatus)
        {
            bool fError = false;
            var itemCounter = 1;
            foreach (var item in SFStatus)
            {
                if (SFStatus.Count > 20)
                {


                    if (itemCounter % 10 == 0)
                    {
                        Console.WriteLine("Processing item {0} of {1}", itemCounter, SFStatus.Count);
                    }
                }
                else
                {
                    Console.WriteLine("Processing item {0} of {1}", itemCounter, SFStatus.Count);

                }
                itemCounter += 1;


                String shelfmark = item.Shelfmark;
                bool validSourceFolder = item.SourceFolderValid;
                bool validAltSourceFolder = item.SourceFolderPathValidElsewhere;

                var ID = item.ID;

               
                if (validSourceFolder)
                {
                string Message = "Valid";
                        Assert.IsTrue(WriteToSharepointColumnByID(spURL, spList, SFCol, shelfmark, ID, Message));
                }
                else if (validAltSourceFolder)
                {
                    string Message = @"Exists with \\ad\collections path";
                    Assert.IsTrue(WriteToSharepointColumnByID(spURL, spList, SFCol, shelfmark, ID, Message));
                }
                else
                {
                    string Message = "Invalid";
                    Assert.IsTrue(WriteToSharepointColumnByID(spURL, spList, SFCol, shelfmark, ID, Message));

                }
                
            }


            return !fError;
        }


        private static String ConstructFullFolderName(string SM, string sF)
        {
            string fullPath = null;
            string SM_folderFormat = SM.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_");

            if (sF.Contains(SM_folderFormat))
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

        private static List<String> GetListOfXMLs(string sF, bool recursive)
        {
            String searchFolder = sF;
            var filters = new String[] { "xml" };
            string[] files = DirectorySearchTools.GetFilesFrom(searchFolder, filters, recursive);
            List<string> listFiles = new List<string>(files);

            return listFiles;
        }

        public static List<List<String>> GetSourceFolderXMLs(List<List<String>> sfStatus, bool recursive)
        {
            List<List<String>> sourceFolderXMLs = new List<List<String>>();
            for (int i = 1; i < sfStatus.Count; i++)
            {
                List<String> item = sfStatus[i];
                var shelfmark = item[1];
                string sourceFolder = "";
                if (string.IsNullOrEmpty(item[6]))
                {
                    sourceFolder = item[2];
                }
                else
                {
                    sourceFolder = item[6];
                }


                List<String> xmlList = new List<String>();
                try
                {
                    xmlList = GetListOfXMLs(sourceFolder, recursive);

                    if (xmlList.Count > 0)
                    {
                    }
                }
                catch
                {
                    xmlList = null;
                }

                sourceFolderXMLs.Add(xmlList);
            }

            return sourceFolderXMLs;
        }

        public static List<String> GetShelfmarkXMLs(String sourceFolder)
        {
            List<String> sourceFolderXMLs = new List<String>();

            DirectoryInfo d = new DirectoryInfo(sourceFolder);
            FileInfo[] Files = d.GetFiles("*.xml*");
            List<String> XMLFiles = Files.Select(x => x.Name).ToList();
            return XMLFiles;
        }
        public static bool IsInvalidFileNameChar(Char c) => c < 64U ?
        (1UL << c & 0xD4008404FFFFFFFFUL) != 0 :
        c == '\\' || c == '|';

        public static bool FilePathHasInvalidChars(string testFilePath)
        {
            bool stringExists = !string.IsNullOrEmpty(testFilePath);
            char[] invalidChars = Path.GetInvalidFileNameChars();
            
            bool hasBadChars = testFilePath.IndexOfAny(invalidChars) >=0;

            return (stringExists && hasBadChars);
        }

        public static List<HMDObject> BadShelfmarkNames(List<HMDObject> itemList)
        {
            List<HMDObject> badShelfmarksIDs = new List<HMDObject>();
            bool protectedCharsFound = false;
            foreach (var item in itemList)
            {
                List<String> flagShelfmark = new List<String>();
                string Shelfmark = item.Shelfmark;

                // Deprecated:
                foreach (char character in Shelfmark)
                {
                   if (IsInvalidFileNameChar(character))
                  {
                      protectedCharsFound = true;

                  }
                }
                // New method:
                //protectedCharsFound = FilePathHasInvalidChars(Shelfmark);
                if (protectedCharsFound)
                {
                    item.BadShelfmark = true;
                }
                badShelfmarksIDs.Add(item);

            }

            return badShelfmarksIDs;
        }

        public static bool CreateSharepointColumn(String SPSite, String SPListName, String newCol)
        {
            bool fError = false;
            bool fieldExists = false;
            try
            {
                using (ClientContext clientContext = new ClientContext(SPSite))
                {
                    // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
                    // This value is NOT List internal name
                    List targetList = clientContext.Web.Lists.GetByTitle(SPListName);
                    clientContext.Load(targetList);
                    clientContext.Load(targetList.Fields);
                    clientContext.ExecuteQuery();
                    for (int i = 0; i < targetList.Fields.Count; i++)
                    {
                        if (targetList.Fields[i].Title == newCol)
                        {
                            fieldExists = true;
                        }
                    }
                }
                if (!fieldExists)
                {
                    using (ClientContext clientContext = new ClientContext(SPSite))
                    {
                        // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
                        // This value is NOT List internal name
                        List targetList = clientContext.Web.Lists.GetByTitle(SPListName);
                        FieldCollection collField = targetList.Fields;

                        string fieldSchema = "<Field Type='Text' DisplayName='" + newCol + "' Name='" + newCol + "' />";
                        collField.AddFieldAsXml(fieldSchema, true, AddFieldOptions.AddToDefaultContentType);

                        clientContext.Load(collField);
                        clientContext.ExecuteQuery();

                    }
                }
                //else // this just isn't needed unless debugging
                //{
                //    Console.WriteLine("The SharePoint column you're trying to add already exists!");
                //}
                return !fError;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing new column {0}. Exception: {1}", newCol, ex);
                fError = true;
                return fError;
            }
        }

        // Will delete a sharepoint column - use this sparingly!
        public static bool DeleteSharepointColumn(String SPSite, String SPListName, String delCol)
        {
            bool fError = false;
            bool fieldExists = false;
            try
            {
                using (ClientContext clientContext = new ClientContext(SPSite))
                {
                    // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
                    // This value is NOT List internal name
                    List targetList = clientContext.Web.Lists.GetByTitle(SPListName);
                    clientContext.Load(targetList);
                    clientContext.Load(targetList.Fields);
                    clientContext.ExecuteQuery();
                    for (int i = 0; i < targetList.Fields.Count; i++)
                    {
                        if (targetList.Fields[i].Title == delCol)
                        {
                            fieldExists = true;
                        }
                    }
                }
                if (fieldExists)
                {
                    using (ClientContext clientContext = new ClientContext(SPSite))
                    {
                        // This value is NOT List internal name
                        List targetList = clientContext.Web.Lists.GetByTitle(SPListName);

                        // Get field from site collection using internal name or display name
                        Field oField = clientContext.Web.AvailableFields.GetByInternalNameOrTitle(delCol);

                        // Delete field
                        oField.DeleteObject();

                        clientContext.ExecuteQuery();

                    }
                }
                else
                {
                    Console.WriteLine("The SharePoint column you're trying to delete doesn't exist!");
                }
                return !fError;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error deleting column {0}. Exception: {1}", delCol, ex);
                fError = true;
                return fError;
            }
        }

        // Add writing functionality here...
       
        public static bool WriteToSharepointColumnByID(String SPSite, String SPListName, String writeCol, String shelfmark, string ID, String Message)
        {
          

            bool fError = false;
            using (ClientContext clientContext = new ClientContext(SPSite))
            {
                try
                {
                    var theID = ID;
                    List targetList = clientContext.Web.Lists.GetByTitle(SPListName);
                    clientContext.Load(targetList);
                    SP.ListItem item = targetList.GetItemById(theID);
                    clientContext.Load(item); // loading all the fields
                    clientContext.ExecuteQuery();

                   
                    if (item.FieldValues.ContainsKey(writeCol))
                    {
                        var currentColumnValue = item[writeCol];
                        item[writeCol] = Message;
                        item.Update();
                        clientContext.ExecuteQuery(); // commits any changes to the sharepoint site
                    }
                    else
                    {
                        Console.WriteLine("The sharepoint column you're trying to write to does not exist!");
                    }
                    

                }

                catch (Exception ex)
                {
                    Console.WriteLine("Error writing to Sharepoint for shelfmark {0}. Exception: {1}",shelfmark, ex);
                    fError = true;
                    return !fError;
                }
            }

            return !fError;

        }
    }
    }


    



using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
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
            catch (Exception ex)
            {
                Console.WriteLine("Failed to find sharepoint site. " + ex.Message);

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
                    listNames.Add(oList.Title);
                }
                return listNames;
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
            catch
            {
                return null;
            }
        }



        public static List<List<String>> GetSharePointListFieldContents(string sURL, string lName, string env, string inputvar)
        {
            var myID = "";
            var myTitle = "";
            var myLoc = "";
            if (env == "test")
            {
                try
                {

                    ClientContext clientContext = new ClientContext(sURL);
                    SP.List oList = clientContext.Web.Lists.GetByTitle(lName);


                    CamlQuery camlQuery = new CamlQuery();
                    var myQuery = @"<View><Query><Where><Contains>" + inputvar + @"</Contains></Where></Query></View>";
                    camlQuery.ViewXml = String.Format(myQuery);
                    ListItemCollection oItems = oList.GetItems(camlQuery);

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

                        catch(Exception ex)
                        {
                            Console.WriteLine("No source folder found! Exception {0}", ex);
                            continue; // If the itemLocation is empty, we don't care, but this throws an exception so need to skip over this item
                        }
                        if (itemLocation != null)
                        {
                            String rowString = String.Format("ID: {0} \t Project: {1} \t Title: {2} \t Location: {3}", itemID, oListItem.FieldValues["Project_x0020_Name"].ToString(), itemTitle, itemLocation);
                            myID = itemID;
                            myTitle = itemTitle;
                            myLoc = itemLocation;
                            Console.WriteLine(rowString);
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
            else if (env == "prod")
            {

                try
                {

                    ClientContext clientContext = new ClientContext(sURL);
                    SP.List oList = clientContext.Web.Lists.GetByTitle(lName);

                    CamlQuery camlQuery = new CamlQuery();

                    var myQuery = @"<View><Query><Where><Contains>" + inputvar + @"</Contains></Where></Query></View>";
                    camlQuery.ViewXml = String.Format(myQuery);
                    SP.ListItemCollection oItems = oList.GetItems(camlQuery);

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
                            Console.WriteLine(rowString);
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
            else
            {
                Console.WriteLine("You forgot to set the environment.");
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
            fESTitles.Add("Full Source Folder Path");

            folderExistenceStatus.Add(fESTitles); // Add the titles list as the first item in the list of lists

            foreach (var item in itemList)
            {
                List<String> itemStatus = new List<String>();
                bool sourceFolderValid = false;
                bool sourceFolderValidElsewhere = false;

                string fullSourceFolderPath = "";

                //Console.WriteLine("{0} \t {1} \t {2}", item[0], item[1],item[2]);
                string ID = item[0];
                string Shelfmark = item[1];
                string sourceFolderSP = item[2];
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
                    itemStatus.Add(sourceFolderValid.ToString());
                    itemStatus.Add(fullSourceFolderPath);
                    itemStatus.Add(sourceFolderValidElsewhere.ToString());




                }
                catch
                {
                    bool DirectoryStatus = false;
                    bool altDirectoryStatus = false;
                    itemStatus.Add(ID);
                    itemStatus.Add(Shelfmark);
                    itemStatus.Add(sourceFolder);
                    itemStatus.Add(DirectoryStatus.ToString());
                    itemStatus.Add(altDirectoryStatus.ToString());
                    itemStatus.Add(sourceFolderValid.ToString());
                    itemStatus.Add(fullSourceFolderPath);
                    itemStatus.Add(sourceFolderValidElsewhere.ToString());



                    // really need to handle this exception properly!
                }
                folderExistenceStatus.Add(itemStatus);


                // Need to decide how to do reporting with this - final bool shows whether the output folder needs flagging.
                // Source folder needs flagging if: source folder not found at all, or source folder found under \\ad\collections but not under the location given


            }
            return folderExistenceStatus;

        }

        public static bool ReportSourceFolderStatus(string spURL, string spList, string SFCol, List<List<String>> SFStatus)
        {
            bool fError = false;
            for (int i = 1; i < SFStatus.Count; i++)
            {
                var item = SFStatus[i];
                String shelfmark = item[1].ToString();
                String validSourceFolder = item[5].ToString();
                String validAltSourceFolder = item[7].ToString();

                //Int32 ID = Int32.Parse(item[0]);
                var ID = item[0];

                if (!String.IsNullOrEmpty(validSourceFolder))
                {
                    if (validSourceFolder.ToUpper().ToLower() == "true")
                    {
                        string Message = "Valid";
                        Assert.IsTrue(WriteToSharepointColumnByID(spURL, spList, SFCol, shelfmark, ID, Message));
                    }
                    else if (validAltSourceFolder.ToUpper().ToLower() == "true")
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

        public static List<List<String>> BadShelfmarkNames(List<List<String>> itemList)
        {
            List<List<String>> badShelfmarksIDs = new List<List<String>>();
            bool protectedCharsFound = false;
            foreach (var item in itemList)
            {
                List<String> flagShelfmark = new List<String>();
                string Shelfmark = item[1];
                foreach (char character in Shelfmark)
                {
                    if (IsInvalidFileNameChar(character))
                    {
                        protectedCharsFound = true;

                    }
                }
                if (protectedCharsFound)
                {
                    flagShelfmark.Add(item[0]);
                    flagShelfmark.Add(Shelfmark);
                    badShelfmarksIDs.Add(flagShelfmark);
                }
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
                else
                {
                    Console.WriteLine("The SharePoint column you're trying to add already exists!");
                }
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
                        // clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
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
        public static bool WriteToSharepointColumnByShelfmark(String SPSite, String SPListName, String writeCol, List<String> shelfmarks)
        {
            bool fError = false;

            try
            {

                ClientContext ctx = new ClientContext(SPSite);
                List list = ctx.Web.Lists.GetByTitle(SPListName);

                foreach (String shelfmark in shelfmarks)
                {

                    CamlQuery camlQuery = new CamlQuery();
                    camlQuery.ViewXml = "<Where><Contains><FieldRef Name ='Title'/><Value Type='Text'>" + shelfmark + "</ Value ></ Contains ></ Where > ";
                    ListItemCollection items = list.GetItems(camlQuery);
                    ctx.Load(items); // loading all the fields
                    ctx.ExecuteQuery();

                    foreach (var item in items)
                    {
                        if (item["Shelfmark"].ToString() == shelfmark) // really need to make sure this is the right shelfmark!
                        {
                            item[writeCol] = "BadCharacters";
                            item.Update(); // remember changes
                            ctx.ExecuteQuery(); // commit changes to the server
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing to Sharepoint. Exception: {0}", ex);
                fError = true;
                return !fError;
            }

            return !fError;


        }
        public static bool WriteToSharepointColumnBySingleShelfmark(String SPSite, String SPListName, String writeCol, String shelfmark, String Message)
        {
            bool fError = false;
            using (ClientContext clientContext = new ClientContext(SPSite))
            {
                try
                {
                    List targetList = clientContext.Web.Lists.GetByTitle(SPListName);
                    clientContext.Load(targetList);
                    //CamlQuery camlQuery = new CamlQuery();
                    //camlQuery.ViewXml = "<View><Query><Where><Contains><FieldRef Name ='Title'/><Value Type='Text'>" + shelfmark + "</ Value ></ Contains ></ Where ></Query></View> ";
                    CamlQuery camlQuery = CamlQuery.CreateAllItemsQuery();
                    ListItemCollection items = targetList.GetItems(camlQuery);
                    clientContext.Load(items); // loading all the fields
                    clientContext.ExecuteQuery();

                    foreach (var item in items)
                    {
                        if (item.FieldValues["Title"].ToString()== shelfmark) // really need to make sure this is the right shelfmark!
                        {
                            item.FieldValues[writeCol] = Message;
                            item.Update(); // remember changes
                            clientContext.ExecuteQuery(); // commit changes to the server
                        }
                    }

                }

                catch (Exception ex)
                {
                    Console.WriteLine("Error writing to Sharepoint. Exception: {0}", ex);
                    fError = true;
                    return !fError;
                }
            }

                return !fError;

            }
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

                    if (item.FieldValues["Title"].ToString() == shelfmark) // really need to make sure this is the right shelfmark!
                    {
                        if (item.FieldValues.ContainsKey(writeCol))
                        {
                        var currentColumnValue = item[writeCol];
                        item[writeCol] = Message;
                        item.Update(); // remember changes
                        clientContext.ExecuteQuery(); // commit changes to the server
                        }
                        else
                        {
                            Console.WriteLine("The sharepoint column you're trying to write to does not exist!");
                        }
                    }

                    // For debugging - delete before pushing commit
                  /*
                        // Check things have written!
                        List resultList = clientContext.Web.Lists.GetByTitle(SPListName);
                        clientContext.Load(resultList);
                        SP.ListItem resultItem = targetList.GetItemById(theID);
                        clientContext.Load(resultItem); // loading all the fields
                        clientContext.ExecuteQuery();
                        var updatedField = item.FieldValues[writeCol];
                 */
                    

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


    



using System;
using System.IO;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Linq;



namespace HMDSharepointChecker
{
    public class LibraryAPIs
    {
        public class IamsItem // adding additional information to the things retrieved from the HMD sharepoint site
        {
            public string SharepointID { get; set; }
            public string ItemShelfmark { get; set; }
            public string ItemDescription { get; set; }
            public string ArkIdentifier { get; set; }
            public string SubSubSeries { get; set; }
            public string LogicalLabel { get; set; }
            public string LogicalType { get; set; }
            public List<string> ChildRecordTitles { get; set; }
            public List<string> DeletedChildRecordTitles { get; set; }
            public bool DeleteFlagPresent { get; set; }
            public string ItemType { get; set; }
            public string CatalogueStatus { get; set; }
         



            public IamsItem(string sharepointID, string itemShelfmark, string iamsTitle, string iamsArk, string subSubSeries, string logicalLabel, string logicalType, List<string> childRecords, List<string> deletedChildRecordTitles, bool deleteFlag, string itemType, string catStatus)
            {
                SharepointID = sharepointID;
                ItemShelfmark = itemShelfmark;
                ItemDescription = iamsTitle;
                ArkIdentifier = iamsArk;
                SubSubSeries = subSubSeries;
                LogicalLabel = logicalLabel;
                LogicalType = logicalType;
                ChildRecordTitles = childRecords;
                DeletedChildRecordTitles = deletedChildRecordTitles;
                DeleteFlagPresent = deleteFlag;
                ItemType = itemType;
                CatalogueStatus = catStatus;

            }
            public IamsItem()
            {
                SharepointID = null;
                ItemShelfmark = null;
                ItemDescription = null;
                ArkIdentifier = null;
                SubSubSeries = null;
                LogicalLabel = null;
                LogicalType = null;
                ChildRecordTitles = null;
                DeletedChildRecordTitles = null;
                DeleteFlagPresent = false;
                ItemType = null;
                CatalogueStatus = null;
            }
        }

        public class AlephItem // adding additional information to the things retrieved from the HMD sharepoint site
        {
            public string FieldTitle { get; set; }
            public string FieldValue { get; set; }


            public AlephItem(string fieldTitle, string fieldValue)
            {
                FieldTitle = fieldTitle;
                FieldValue = fieldValue;

            }
            public AlephItem()
            {
                FieldTitle = null;
                FieldValue = null;
            }
        }



        static public readonly string IAMSURL = @"http://v12l-iams3/IAMSRestAPILive/api/archive/GetRecordByreference?reference=";

        public static List<IamsItem> queryMetadataAPIs (String spURL,String spList,List<HMDObject> itemList, bool writeToSharepoint)
        {
            // This function makes sure that you aren't querying Aleph for manuscripts, or similar
            List<IamsItem> IAMSRecords = new List<IamsItem>();
            foreach (HMDObject item in itemList)
            {
                if(item.MetadataSource.ToUpper().ToLower().Contains("aleph"))
                {
                    List<AlephItem> returnedAlephRecords = GetAlephRecords(item);
                    string outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    outFolder += @"\HMDSharepoint_AlephRecords\";
                    outFolder += item.Shelfmark+"_"+item.SystemNumber;
                    if (!Directory.Exists(outFolder))
                    {
                        Directory.CreateDirectory(outFolder);
                    }
                    writeAlephCSV(item.Shelfmark, returnedAlephRecords, outFolder);                    

                }
                else if (item.MetadataSource.ToLower().ToUpper().Contains("IAMS"))
                {
                    var IAMSitem = GetIAMSRecords(item);
                    IAMSRecords.Add(IAMSitem);

                }

            }
            foreach(var iamsItem in IAMSRecords)
            {
                if (iamsItem.SharepointID != null)
                {
                    try
                    {
                        String columnName = "IAMSCatalogueStatusIsPublished";
                        var theShelfmark = iamsItem.ItemShelfmark;
                        var itemID = iamsItem.SharepointID;
                        var message = "";
                        if (iamsItem.CatalogueStatus.ToUpper().ToLower().Contains("published"))
                        {
                            message = "Yes";
                        }
                        else
                        {
                            message = "No";
                        }
                        if (writeToSharepoint)
                        {
                            SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                            SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                        }

                        if(iamsItem.ChildRecordTitles.Count > 0)
                        {
                            // write IAMS child record titles to a CSV
                            string outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            outFolder += @"\HMDSharepoint_IamsChildRecords\";
                            outFolder += iamsItem.ItemShelfmark;
                            if (!Directory.Exists(outFolder))
                            {
                                Directory.CreateDirectory(outFolder);
                            }
                            writeIAMSCSV(iamsItem, outFolder);

                        }
                    }

                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing IAMS catalogue status. \nException {0}", ex);
                    }
                    try
                    {
                        String columnName = "Shelfmark_HasIAMSDeleteFlag";
                        var theShelfmark = iamsItem.ItemShelfmark;
                        var itemID = iamsItem.SharepointID;
                        var message = "";
                        if (iamsItem.DeleteFlagPresent)
                        {
                            message = "Yes";
                        }
                        else
                        {
                            message = "No";
                        }
                        if (writeToSharepoint)
                        {

                            SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                            SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing IAMS delete flag status. \nException {0}", ex);

                    }
                }
                else
                {
                    continue;
                }
                
            }

                return IAMSRecords;
        }

        public static IamsItem GetIAMSRecords(HMDObject item)
        {

            // Remember to get the list of child records and pass this back to the image order csv somehow...
            String shelfmark = item.Shelfmark;
            String itemID = null;

            if (item.ID != null ) // allows for some flexibility on being tied to sharepoint here
            {
                itemID = item.ID;
            }
           

            // build the request
            // Need to use a request with the 'None' format to get the catalogue publication status/state
            // Also need 'itemtype' field where this can be fonds etc

            // The only output format of the IAMS request that contains the catalogue publication status is "NONE"
            // So make two requests to get all the information required

            var iamsCatStatusRequest = IAMSURL + shelfmark.Replace(" ", "%20").Replace(",", "%2C") + @"&format=None";
            var iamsGetUrl = IAMSURL + shelfmark.Replace(" ", "%20").Replace(",", "%2C") + @"&format=Qatar"; // Could use DIPS format, but Qatar gets you more information


            // Run the first request:
                var xmlTextReaderIamsNone = new XmlTextReader(iamsCatStatusRequest);
                var xmlDocumentIamsNone = new XmlDocument();
                try
                {
                    xmlDocumentIamsNone.Load(xmlTextReaderIamsNone);
                }
                catch
                {
                    return null;
                }
                var xmlNodeItemType= xmlDocumentIamsNone.SelectSingleNode("//Header//RecordDetails//ItemType");
                var xmlNodeCatStatus = xmlDocumentIamsNone.SelectSingleNode("//Header//RecordDetails//Status");
                


                var iamsItemType = string.Empty;
                if (xmlNodeItemType != null) iamsItemType = xmlNodeItemType.InnerText;

                var iamsCatStatus = string.Empty;
                if (xmlNodeCatStatus != null) iamsCatStatus = xmlNodeCatStatus.InnerText;



                var xmlTextReaderIams = new XmlTextReader(iamsGetUrl);
                var xmlDocumentIams = new XmlDocument();
                try
                {
                    xmlDocumentIams.Load(xmlTextReaderIams);
                }
                catch
                {
                    return null;
                }
                var xmlNodeArk = xmlDocumentIams.SelectSingleNode("//MDARK");
                var xmlNodeTitle = xmlDocumentIams.SelectSingleNode("//Title");
                var xmlNodeSubSubSeries = xmlDocumentIams.SelectSingleNode("//Ancestors//Ancestor[@level='1']//Reference");
                var xmlLogicalLabel = xmlDocumentIams.SelectSingleNode("//LogicalLabel");
                var xmlLogicalType = xmlDocumentIams.SelectSingleNode("//LogicalType");
                var xmlChildRecords = xmlDocumentIams.SelectNodes("//Children//Child//Reference");
                var xmlShelfmark = xmlDocumentIams.SelectSingleNode("//Reference");
            

                var logicalLabel = string.Empty;
                var logicalType = string.Empty;

                if (xmlNodeArk == null) return null;
                var iamsArk = xmlNodeArk.InnerText;

                var iamsRetrievedShelfmark = string.Empty;
                bool deleteFlagShelfmark = false;
                if (xmlShelfmark != null) iamsRetrievedShelfmark = xmlShelfmark.InnerText;
                if (iamsRetrievedShelfmark.ToString().Contains("DEL") || iamsRetrievedShelfmark.ToString().Contains(@"D/"))
            {
                deleteFlagShelfmark = true;
            }

                var iamsTitle = string.Empty;
                if (xmlNodeTitle != null) iamsTitle = xmlNodeTitle.InnerText;

                var subSubSeries = string.Empty;
                if (xmlNodeSubSubSeries != null) subSubSeries = xmlNodeSubSubSeries.InnerText;

                if (xmlLogicalLabel != null)
                    logicalLabel = xmlLogicalLabel.InnerText;

                if (xmlLogicalType != null)
                    logicalType = xmlLogicalType.InnerText;

                List<string> childRecordTitles = new List<String>();
                List<string> deletedChildRecordTitles = new List<String>();

                bool containsDeletedChildRecords = false;
                if (xmlChildRecords != null)
                {
                
                foreach (XmlNode record in xmlChildRecords)
                {
                    var innerText = record.InnerText;
                    childRecordTitles.Add(record.InnerText);
                    if(record.InnerText.ToString().Contains("DEL") || record.InnerText.ToString().Contains(@"D/"))
                    {
                        containsDeletedChildRecords = true;
                        deletedChildRecordTitles.Add(record.InnerText);
                    }
                }


                }
            bool isDeleteFlagPresent = false;
            if(containsDeletedChildRecords || deleteFlagShelfmark)
            {
                isDeleteFlagPresent = true;
                
            }

            var iamsItem = new IamsItem
            {
                SharepointID = itemID,
                ItemShelfmark = shelfmark,
                ItemDescription = iamsTitle,
                ArkIdentifier = iamsArk,
                SubSubSeries = subSubSeries,
                LogicalLabel = logicalLabel,
                LogicalType = logicalType,
                ChildRecordTitles = childRecordTitles,
                DeletedChildRecordTitles = deletedChildRecordTitles,
                DeleteFlagPresent = isDeleteFlagPresent,
                ItemType = iamsItemType,
                CatalogueStatus = iamsCatStatus
                };

            

                return iamsItem;
        }

        public static List<AlephItem> GetAlephRecords(HMDObject item)
        {
            String shelfmark = item.Shelfmark;
            String itemID = null;

            if (item.ID != null) // allows for some flexibility on being tied to sharepoint here
            {
                itemID = item.ID;
            }

            if(item.SystemNumber.Length <1)
            {
                Console.WriteLine("No Aleph system number found when attempting to look up Aleph item for shelfmark: {0}", shelfmark);
                return null;
            }
            var systemNumber = item.SystemNumber;


            // build the request
            var alephURL = "http://xserver.bl.uk/X";
            var alephRequest = alephURL+"?op=find_doc&doc_num="+systemNumber+ "&base=BLL01"; // uses default BL base system number
          
            // Run the first request:
            var xmlTextReaderAleph = new XmlTextReader(alephRequest);
            var xmlDocumentAleph = new XmlDocument();
            try
            {
                xmlDocumentAleph.Load(xmlTextReaderAleph);
            }
            catch
            {
                return null;
            }
            var alephRecords = xmlDocumentAleph.SelectNodes("//find-doc//record//metadata//oai_marc");


            foreach (XmlNode element in alephRecords)

            {
                var thisvar = element.InnerText;
            }


            // this needs sorting out, but basically want a csv to write the attribute ID as a column and the value as a field
            // not sure what to do yet if the node has sub-lists, but debug this step by step and see if the node type changes?
            // nodes with sublists are all varfields!

            List<AlephItem> alephReturnedItems = new List<AlephItem>();

            var doc = XDocument.Load(alephRequest);


            var fixedfields = from @fixedfield in doc.Descendants("fixfield")
                              let fixedfieldName = (string)@fixedfield.Attribute("id")
                              let fixedfieldValue = (string)fixedfield.Value
                            select new
                            {
                                FixedFieldName = fixedfieldName,
                                FixedFieldValue = fixedfieldValue
                            };

            var varfields = from @varfield in doc.Descendants("varfield")
                        let varfieldName = (string)@varfield.Attribute("id")
                        from subfield in @varfield.Descendants("subfield")
                        select new
                        {
                            VarfieldName = varfieldName,
                            SubfieldLabel = (string)subfield.Attribute("label"),
                            SubfieldContents = (string)subfield.Value,
                        };


            foreach( var field in fixedfields)
            {
                var alephItem = new AlephItem();
                alephItem.FieldTitle = field.FixedFieldName;
                alephItem.FieldValue = field.FixedFieldValue;
                alephReturnedItems.Add(alephItem);
            }
            foreach (var field in varfields)
            {
                var alephItem = new AlephItem();
   
                alephItem.FieldTitle= field.VarfieldName+field.SubfieldLabel;
                alephItem.FieldValue = field.SubfieldContents;
                alephReturnedItems.Add(alephItem);
            }

            return alephReturnedItems;
        }

        private static bool writeAlephCSV(string shelfmark,List<AlephItem> AlephRecords, String outFolder)
        {
            bool fError = false;

            try // to write the csv...
            {
                List<String> strHeaders = new List<string>();
                foreach(AlephItem aItem in AlephRecords)
                {
                    strHeaders.Add(aItem.FieldTitle);
                }

                System.Text.UnicodeEncoding uce = new System.Text.UnicodeEncoding();
                System.Text.Encoding utf8 = System.Text.Encoding.UTF8; // changed from uce to utf8

                string fNameString = "AlephRecords";
                string outPath = outFolder + @"\" + fNameString + ".csv";

                if (!File.Exists(outPath)) // only write once per aleph system number...
                {
                    /*
                    String lastModified = File.GetLastWriteTime(outPath).ToString("yyyyMMdd_HH-mm-ss");
                    string oldFilePath = outFolder + @"\" + fNameString + "_" + lastModified + ".csv";

                    try
                    {
                        File.Move(outPath, oldFilePath); // moves existing imageorder.csv file to include the last modified time
                                                         // leaves newest created file as 'imageorder.csv'
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Could not move existing AlephRecords.csv from {0} to {1}.\nException: {2}", outPath, oldFilePath, ex);
                    }

                    // }
                    */


                    using (var sr = new StreamWriter(outPath, false)) // changed from uce to utf8
                    {
                        using (var csvFile = new CsvHelper.CsvWriter(sr, System.Globalization.CultureInfo.InvariantCulture))
                        {
                            csvFile.Configuration.Delimiter = ",";
                            //csvFile.Configuration.HasExcelSeparator = true;

                            foreach (var header in strHeaders)
                            {
                                csvFile.WriteField(header);
                            }
                            csvFile.NextRecord(); // skips to next line...
                            var fieldCounter = 0;
                            foreach (var record in AlephRecords)
                            {
                                csvFile.WriteField(record.FieldValue); // field value
                                fieldCounter += 1;
                                var lastRecord = AlephRecords[AlephRecords.Count - 1].FieldValue;
                                if (fieldCounter >= strHeaders.Count)
                                {
                                    csvFile.NextRecord();
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing CSV File: {0}", ex);
                fError = true;
            }
            return !fError;
        }

        private static bool writeIAMSCSV(IamsItem iamsRecord, String outFolder)
        {
            bool fError = false;

            try // to write the csv...
            {
                List<String> strHeaders = new List<string>();
                
                strHeaders.Add("Child shelfmarks");
                

                string fNameString = "IamsRecords";
                string outPath = outFolder + @"\" + fNameString + ".csv";

                if (!File.Exists(outPath)) // only write once per shelfmark...
                {
                    
                    using (var sr = new StreamWriter(outPath, false)) // changed from uce to utf8
                    {
                        using (var csvFile = new CsvHelper.CsvWriter(sr, System.Globalization.CultureInfo.InvariantCulture))
                        {
                            csvFile.Configuration.Delimiter = ",";
                            //csvFile.Configuration.HasExcelSeparator = true;

                            csvFile.WriteField("Shelfmark");
                            csvFile.WriteField("Item type");
                            csvFile.WriteField("Catalogue status");
                            csvFile.NextRecord(); // skips to next line...

                            csvFile.WriteField(iamsRecord.ItemShelfmark);
                            csvFile.WriteField(iamsRecord.ItemType);
                            csvFile.WriteField(iamsRecord.CatalogueStatus);
                            csvFile.NextRecord(); // skips to next line...




                            foreach (var header in strHeaders)
                            {
                                csvFile.WriteField(header);
                            }
                            csvFile.NextRecord(); // skips to next line...
                            foreach (var cShelfmark in iamsRecord.ChildRecordTitles)
                            {
                                csvFile.WriteField(cShelfmark); // field value
                                csvFile.NextRecord();
                             }
                            
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error writing IAMS CSV File: {0}", ex);
                fError = true;
            }
            return !fError;
        }


    }
}



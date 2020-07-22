using System;
using System.Collections.Generic;
using System.Xml;


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



            public AlephItem(string sharepointID, string itemShelfmark, string iamsTitle, string iamsArk, string subSubSeries, string logicalLabel, string logicalType, List<string> childRecords, List<string> deletedChildRecordTitles, bool deleteFlag)
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

            }
            public AlephItem()
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
            }
        }



        static public readonly string IAMSURL = @"http://v12l-iams3/IAMSRestAPILive/api/archive/GetRecordByreference?reference=";

        public static bool queryMetadataAPIs (List<HMDObject> itemList)
        {
            // This function makes sure that you aren't querying Aleph for manuscripts, or similar
            List<IamsItem> IAMSRecords = new List<IamsItem>();

            bool ferror = false;
            foreach (HMDObject item in itemList)
            {
                if(item.MetadataSource.ToUpper().ToLower().Contains("aleph"))
                {
                    continue; // currently do nothing for aleph, fucntionality still WIP

                }
                else if (item.MetadataSource.ToLower().ToUpper().Contains("IAMS"))
                {
                    var IAMSitem = GetIAMSRecords(item);
                    IAMSRecords.Add(IAMSitem);
                }

            }

                return !ferror;
        }

        public static IamsItem GetIAMSRecords(HMDObject item)
        {
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

                var logicalLabel = string.Empty;
                var logicalType = string.Empty;

                if (xmlNodeArk == null) return null;
                var iamsArk = xmlNodeArk.InnerText;

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
                DeleteFlagPresent = containsDeletedChildRecords,
                ItemType = iamsItemType,
                CatalogueStatus = iamsCatStatus
                };

            

                return iamsItem;
        }



    }
}



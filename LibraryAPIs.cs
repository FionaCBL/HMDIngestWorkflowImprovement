using System;
using System.Collections.Generic;
using System.Xml;


namespace HMDSharepointChecker
{
    class LibraryAPIs
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
            public string ChildRecords { get; set; }



            public IamsItem(string sharepointID, string itemShelfmark, string iamsTitle, string iamsArk, string subSubSeries, string logicalLabel, string logicalType, string childRecords)
            {
                SharepointID = sharepointID;
                ItemShelfmark = itemShelfmark;
                ItemDescription = iamsTitle;
                ArkIdentifier = iamsArk;
                SubSubSeries = subSubSeries;
                LogicalLabel = logicalLabel;
                LogicalType = logicalType;
                ChildRecords = childRecords;

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
                ChildRecords = null;
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
                String itemID = item.ID;
                // build the request
                var iamsGetUrl = IAMSURL + shelfmark.Replace(" ", "%20").Replace(",", "%2C") + @"&format=Qatar"; // Could use DIPS format, but Qatar gets you more information
                var xmlTextReaderIams = new XmlTextReader(iamsGetUrl);
                var xmlDocumentIams = new XmlDocument();
                xmlDocumentIams.Load(xmlTextReaderIams);

                var xmlNodeArk = xmlDocumentIams.SelectSingleNode("//MDARK");
                var xmlNodeTitle = xmlDocumentIams.SelectSingleNode("//Title");
                var xmlNodeSubSubSeries = xmlDocumentIams.SelectSingleNode("//Ancestors//Ancestor[@level='1']//Reference");
                var xmlLogicalLabel = xmlDocumentIams.SelectSingleNode("//LogicalLabel");
                var xmlLogicalType = xmlDocumentIams.SelectSingleNode("//LogicalType");
                var xmlChildRecords = xmlDocumentIams.SelectSingleNode("//Children");

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

                var iamsChildRecords = string.Empty;
                if (xmlChildRecords != null) iamsChildRecords = xmlChildRecords.InnerText; 

                var iamsItem = new IamsItem
                {
                    SharepointID = itemID,
                    ItemShelfmark = shelfmark,
                    ItemDescription = iamsTitle,
                    ArkIdentifier = iamsArk,
                    SubSubSeries = subSubSeries,
                    LogicalLabel = logicalLabel,
                    LogicalType = logicalType,
                    ChildRecords = iamsChildRecords
                };

            

                return iamsItem;
        }


    }
}



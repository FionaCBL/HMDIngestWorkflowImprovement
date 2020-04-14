using System;
using System.Web;
using System.DirectoryServices.AccountManagement;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;        
using System.Net;
using Microsoft.VisualStudio.TestTools.UnitTesting;


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
            Assert.IsTrue(DigitisationWorkflowTitles.Count != 0);
            
            
            
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

                    /*
                    View view = oList.Views.GetByTitle("All Items");
                    clientContext.Load(view);
                    clientContext.ExecuteQuery();


                    // Create a new CAML query object and store the query from the custom view
                    CamlQuery query = new CamlQuery();
                    query.ViewXml = view.ViewQuery;

                    // Based on the query load the items


                    ListItemCollection items = oList.GetItems(query);
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();
                    Console.Write(items.Count);
                    var noItems = items.Count;

                    List<string> listItems = new List<string>();

                    foreach (var item in items)
                    {
                        var thisItemTitle = item["Title"];
                        listItems.Add(item["Title"].ToString());
                    }
                    return listItems;
                    */

                    /// first thing I commented out

                    /*
                    // use undefined camlQuery to get all list items
                    ListItemCollection collListItem = oList.GetItems(camlQuery);
                    clientContext.Load(collListItem);
                    clientContext.ExecuteQuery();

                    List<string> listItems = new List<string>();
                    foreach (ListItem oListItem in collListItem)
                    {

                        Console.WriteLine("ID: {0} \nTitle: {1} \nBody: {2}", oListItem.Id, oListItem["Title"], oListItem["Body"]);
                        listItems.Add(oListItem["Title"].ToString());
                    }
                    return listItems;

                    */
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


    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    public class FileLabels
    {


        public string FileName { get; set; }
        public string FlagStatus { get; set; }
        public string ObjectType { get; set; }
        public string Label { get; set; }
        public string OrderNumber { get; set; }
        public int SubOrder { get; set; }


        public FileLabels(string filename, string flagstatus, string objtype, string label, string ordernum, int subOrder)
        {
            FileName = filename;
            FlagStatus = flagstatus;
            ObjectType = objtype;
            Label = label;
            OrderNumber = ordernum;
            SubOrder = subOrder;

        }
        public FileLabels()
        {
            FileName = null;
            FlagStatus = null;
            ObjectType = null;
            Label = null;
            OrderNumber = null;
            SubOrder = -999;
        }

    }
    class InputOrderSpreadsheetTools
    {
        public static List<List<FileLabels>> listAllShelfmarkFilesTIFXML(List<HMDObject> sharepointOut, String env, String spURL, String spList)
        {
            List<List<FileLabels>> allShelfmarkTIFAndLabels = new List<List<FileLabels>>();
            Console.WriteLine("=======================================\nGenerating image order csv and performing ALTOXML checks...\n=======================================");

            var thisItem = 1;
            foreach (var item in sharepointOut)
            {
                if (sharepointOut.Count > 20)
                {


                    if (thisItem % 10 == 0)
                    {
                        Console.WriteLine("Processing item {0} of {1}", thisItem,sharepointOut.Count);
                    }
                }
                else
                {
                    Console.WriteLine("Processing item {0} of {1}", thisItem, sharepointOut.Count);

                }
                thisItem += 1;


                List<String> shelfmarkTIFs = new List<String>();
                List<FileLabels> shelfmarkLabels = new List<FileLabels>();
                bool validPath = false;
                var itemID = item.ID;
                var shelfmark = item.Shelfmark;

                if (item.SourceFolderValid)
                {
                    string sourceFolder = "";

                    if (string.IsNullOrEmpty(item.FullSourceFolderPath))
                    {
                        sourceFolder = item.SourceFolderPath;
                    }
                    else
                    {
                        sourceFolder = item.FullSourceFolderPath;
                    }

                    // Once you've got sourceFolder, need to get into the actual image folders...

                    // this is a bit of a mess at the moment, sort this out
                    var tifFolder = "";
                    sourceFolder = sourceFolder.TrimEnd('\\');
                    sourceFolder = sourceFolder.ToLower();

                    //Console.WriteLine("Source Folder: {0}", sourceFolder);
                    try
                    {
                        if (sourceFolder.ToUpper().ToLower().Contains("tif"))
                        {
                            tifFolder = sourceFolder;
                        }
                        else
                        {

                            var subFolders = Directory.GetDirectories(sourceFolder);
                            if (subFolders.Length > 0)
                            {

                                foreach (var subFolder in subFolders)
                                {
                                    //Console.WriteLine("Testing subFolder: {0}", subFolder);
                                    if (subFolder.ToUpper().ToLower().Contains("tif"))
                                    {
                                        tifFolder = subFolder;
                                        //   Console.WriteLine("Found subfolder for folder {0}", sourceFolder);
                                    }
                                }
                            }
                            else
                            { // this case rarely gets called, but needs to be accounted for
                                DirectoryInfo d = new DirectoryInfo(sourceFolder);
                                FileInfo[] Files = d.GetFiles("*.TIF*");
                                if(Files.Length > 0)
                                {
                                    tifFolder = sourceFolder;
                                }
                            }
                        }

                        if (Directory.Exists(tifFolder))
                        {
                            validPath = true;
                        }
                        else
                        {
                            Console.WriteLine("No folder found for shelfmark {0}", shelfmark);
                            validPath = false;
                            // return null;
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Could not find Tif subfolder for sourcefolder {0}. Exception: {1}", sourceFolder, ex);
                        validPath = false;
                        //return null;
                    }
                    // now got the tiff folder, need to check the list of files that appears

                    // first check it exists:
                    if (validPath)
                    {
                        // Get all tifs
                        DirectoryInfo d = new DirectoryInfo(tifFolder);
                        FileInfo[] Files = d.GetFiles("*.TIF*");

                        // Can then add this to a list of strings
                        var numberOfItems = Files.Length; // only do this once per shelfmark
                                                          // do you need this?

                        // to-do: turn the below stuff into a class of its own
                        shelfmarkLabels = mapFileNameToLabels(spURL,spList,shelfmark,itemID,Files, tifFolder);
                        // shelfmarkLabels is a list of FileLabels objects
                        // each FileLabels object corresponds to a single file and contains:
                        // filename
                        // error flag status
                        // object type
                        // image label
                        // order number
                        // (all as strings)


                         if (env == "test") // get this going for prod by sticking it in the actual tifFolder
                        {
                            string outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            outFolder += @"\HMDSharepoint_ImgOrderCSVs\";

                            string SM_folderFormat = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");

                            string folderShelfmark = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                            try
                            {
                                string testTifFolder = string.Empty;
                                var tifFolderItems = tifFolder.Split('\\');

                                if (Regex.Matches(tifFolder, folderShelfmark).Count > 1)

                                {
                                    testTifFolder = tifFolder.Split(new string[] { folderShelfmark }, 2, StringSplitOptions.None)[1];
                                }

                                // If we've got a folder with shelfmark/tiffs then the above line will mess things up
                                // Fix them again with this line
                                //if (testTifFolder.ToUpper().ToLower().Contains("\\tif"))
                                //{
                                //    testTifFolder = "\\" + folderShelfmark;
                                //    testTifFolder += tifFolder.Split(new string[] { folderShelfmark }, 2, StringSplitOptions.None)[1];

                                //}
                                else
                                {
                                    testTifFolder = tifFolderItems[tifFolderItems.Length - 2] + "\\" + tifFolderItems[tifFolderItems.Length - 1];
                                }
                                outFolder += testTifFolder;
                                var last = tifFolderItems[tifFolderItems.Length - 1];
                                var secondLast = tifFolderItems[tifFolderItems.Length - 1];

                                if (!Directory.Exists(outFolder))
                                {
                                    Directory.CreateDirectory(outFolder);
                                }                        // Now write this to a CSV

                                Assert.IsTrue(writeFileLabelsToCSV(shelfmarkLabels, outFolder));
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Couldn't set the proper output ImageOrder.csv folder path. Exception: {0}", ex);
                                outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                outFolder += @"\HMDSharepoint_ImgOrderCSVs" + @"\" + SM_folderFormat;
                                if (!Directory.Exists(outFolder))
                                {
                                    Directory.CreateDirectory(outFolder);
                                }                        // Now write this to a CSV

                                Assert.IsTrue(writeFileLabelsToCSV(shelfmarkLabels, outFolder));
                            }

                        }
                        else if (env == "prod")
                        {
                            // ******  This should write straight to network drives in future ****
                            // Only write to Desktop folder for testing. In future, this functionality will only exist for the test environment.

                            try
                            {
                                
                                    Assert.IsTrue(writeFileLabelsToCSV(shelfmarkLabels, tifFolder));
                        
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error writing ImageOrder.csv for shelfmark {0} in folder {1}.\nException {2}",shelfmark,tifFolder,ex);

                            }

                        }

                        // SORT XML THINGS OUT HERE!
                        bool XMLNumberMismatch = false;
                        bool XMLNamingError = false;
                        bool XMLDipsComplianceError = false;
                        bool XMLVersionError = false;
                        string xmlErrMessage = "";
                        List<string> fileNames = Files.Select(x => x.Name).ToList();
                        List<String> fullNames = Files.Select(x => x.FullName).ToList();

                        List<String> allXMLs = SharepointTools.GetShelfmarkXMLs(tifFolder);
                        List<String> baseFileNames = new List<String>();
                        List<String> baseXMLNames = new List<String>();
                        if (allXMLs.Count > 0) // otherwise don't enter this part
                        {
                            if (allXMLs.Count != numberOfItems)
                            {
                                XMLNumberMismatch = true;
                                xmlErrMessage = "Number of XMLs not equals to number of TIFs";
                            }
                            foreach (String fname in fileNames)
                            {
                                string fnameBase = fname.Substring(0, fname.LastIndexOf('.'));
                                baseFileNames.Add(fnameBase);
                            }
                            foreach (String xmlName in allXMLs)
                            {
                                string xmlnameBase = xmlName.Substring(0, xmlName.LastIndexOf('.'));
                                baseXMLNames.Add(xmlnameBase);
                                String fullXMLPath = tifFolder + @"\" + xmlName;
                                //string lineOne = File.ReadLines(fullXMLPath).First(); // gets the first line from file.
                                String versionNumber = GetXMLVersionNumber(fullXMLPath);
                                float XMLVNum = float.Parse(versionNumber, System.Globalization.CultureInfo.InvariantCulture);
                                if (XMLVNum > 2.0)
                                {
                                    XMLVersionError = true;
                                    xmlErrMessage = "XML Version Error";

                                }
                            }
                            if (!baseFileNames.SequenceEqual(baseXMLNames))
                            {
                                XMLNamingError = true;
                                xmlErrMessage = "XML names not equivalent to TIF names";

                            }

                            if (XMLNumberMismatch | XMLNamingError | XMLDipsComplianceError | XMLVersionError)
                            {
                                if (env == "test")
                                {
                                    Assert.IsTrue(SharepointTools.WriteToSharepointColumnByID(spURL, spList, "ALTOXMLCheck", shelfmark, itemID, xmlErrMessage));
                                }
                                else if (env == "prod")
                                {
                                    Console.WriteLine("Holding off on populting columns in Sharepoint prod version for now");
                                }
                            }
                            else
                            { // nothing wrong, just write an valid status
                                if (env == "test")
                                {
                                    Assert.IsTrue(SharepointTools.WriteToSharepointColumnByID(spURL, spList, "ALTOXMLCheck", shelfmark, itemID, "Valid"));

                                }
                            }



                        }


                    }// if a tif folder is found
                    else // so no tif folder found...
                    {
                        // Think about what's going on here. Ultimately you will write the image order csv to the project folder
                        // If the path is invalid this just isn't going to work
                        // Better to write to sharepoint!

                        // Write to a column "TIF_Folder_Found" or similar!

                        FileLabels errorList = new FileLabels();
                        errorList.FileName = shelfmark;
                        errorList.FlagStatus = "TIF folder not found";
                        continue; // use continue for now, but will need to write out invalid path to a variable at some point
                    }
                } // is source folder valid? 
                else // source folder was never valid
                {
                    FileLabels errorList = new FileLabels();
                    errorList.FileName = shelfmark;
                    errorList.FlagStatus = "Invalid source folder";
                    // Got yourself a shelfmark that needs checking, so obviously things will fail here...
                    continue;
                }
                try
                {
                    allShelfmarkTIFAndLabels.Add(shelfmarkLabels);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: Could not add shelfmark tif information to the overall list of shelfmarks \n Shelfmark: {0} \n Exception: {1}", shelfmark, ex);
                    return null;
                    // this should really never happen, so can leave this in
                }
            } // end of the for loop over each shelfmark


            return allShelfmarkTIFAndLabels;

            // returns you a List<List<FileLabels>>
            // Shelfmark labels are outputted as a list of FileLabels objects - for each shelfmark you will have 
            // a FileLabels object for each file containing variables for: filename, error flag, object type, label and order number
            // so shelfmark labels are a list of FileLabels
            // For all shelfmarks this is then List<List<FileLabels>>
        }

        private static List<FileLabels> mapFileNameToLabels(String spURL, string spList,String inputShelfmark, String itemID, FileInfo[] Files, String tifFolders)
        {

            // Order labels will take a couple of sweeps - one to get front and back matter and then another to do a fine sort of the front and back matter
            List<String> shelfmarkLabels = new List<String>();
            string theShelfmark = "";
            inputShelfmark = inputShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");


            List<string> fileNames = Files.Select(x => x.Name).ToList();

            // Define regular expressions to search for
            // 'initial' versions perform a looser search
            string iFrontMatterReString = @"(.)+((fble)((f)|(fv)|(fr)))\.tif";
            var initialFrontMatterRegex = new Regex(iFrontMatterReString, RegexOptions.IgnoreCase);

            string iFrontFlysheetsReString = @"(.)+(fs[0-9]+(.)+)\.tif";
            var initialFrontFlysheetRegex = new Regex(iFrontFlysheetsReString, RegexOptions.IgnoreCase);

            string iFolioReString = @"(.)+(f)([0-9])+(.)+\.tif";
            var initialFolioRegex = new Regex(iFolioReString, RegexOptions.IgnoreCase);

            string iEndFlysheetsReString = @"(.)+((fse)[0-9]+(.)+)\.tif";
            var initialEndFlysheetsRegex = new Regex(iEndFlysheetsReString, RegexOptions.IgnoreCase);


            string iEndMatterReString = @"(.)+(((fb)((rig)|(rigv)|(rigr)|(spi))))\.tif";
            var initialEndMatterRegex = new Regex(iEndMatterReString, RegexOptions.IgnoreCase);

            string iNumericFolioReString = @"(.)+_([0-9])+\.tif";
            var initialNumericFolioRegex = new Regex(iNumericFolioReString, RegexOptions.IgnoreCase);

            // Sort into front matter, end flysheets, end matter and folios
            // Control shots jut shouldn't get picked up at all by any of these
            // Do need to pick up numerically labelled files though


            // Candidates for each section:
            List<string> cFrontMatter = fileNames.Where(f => initialFrontMatterRegex.IsMatch(f)).ToList();
            List<string> cFrontFlysheets = fileNames.Where(f => initialFrontFlysheetRegex.IsMatch(f)).ToList();
            List<string> cEndFlysheets = fileNames.Where(f => initialEndFlysheetsRegex.IsMatch(f)).ToList();
            List<string> cEndMatter = fileNames.Where(f => initialEndMatterRegex.IsMatch(f)).ToList();
            List<string> cFolios = fileNames.Where(f => initialFolioRegex.IsMatch(f)).ToList();
            List<string> cNumericFolios = fileNames.Where(f => initialNumericFolioRegex.IsMatch(f)).ToList();

            List<string> foundItems = cFrontMatter.Concat(cFrontFlysheets).Concat(cFolios).Concat(cNumericFolios).Concat(cEndFlysheets).Concat(cEndMatter).ToList();


            List<String> otherFiles = (from e in (fileNames.Concat(foundItems))
                                       where !foundItems.Contains(e) select e).ToList(); // get everything not found (should be control shots etc)
                                                                                         // is the above worth doing?


            List<FileLabels> allFilesSorted = new List<FileLabels>(); // this is what you're returning later

            List<FileLabels> frontMatter = new List<FileLabels>();
            List<FileLabels> frontFlysheets = new List<FileLabels>();
            List<FileLabels> endFlysheets = new List<FileLabels>();
            List<FileLabels> endMatter = new List<FileLabels>();
            List<FileLabels> folios = new List<FileLabels>();
            List<FileLabels> numericFiles = new List<FileLabels>();
            List<FileLabels> unclassified = new List<FileLabels>();

            bool dipsCompliant = true;
            bool orderCheck = true;

            string folderDerivedShelfmark = "";
        
            if (cFrontMatter.Any() | cFrontFlysheets.Any() | cFolios.Any() | cEndMatter.Any() | cEndFlysheets.Any())
            {
                // you can be pretty sure its DIPs compliant if you see any titles or any numbered folios
               
                bool FMExists = false;
                bool FFSExists = false;
                bool FOLExists = false;
                bool EFSExists = false;
                bool EMExists = false;
                bool numFOLExists = false;
                if (cFrontMatter.Any())
                {

                    FMExists = true;
                    foreach (string fname in cFrontMatter)

                    {
                        FileLabels frontMatterLabels = new FileLabels();

                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;
                        // replace the string
                        string derivedFilename = split2.Last();
                        string trimmedName = derivedFilename.Trim('f', 's');
                        string noZerosName = trimmedName.TrimStart('0');
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";


                        var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif",@".tif");
                        string matchString = derivedShelfmark+@"_((fble)((f)|(fv)|(fr)))\.tif";
                        var match = Regex.Match(fname,matchString, RegexOptions.IgnoreCase);
                        var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);
                        if (match.Success || altMatch.Success)
                        {
                            frontMatterLabels.FileName = fname;
                            var fblef = Regex.Match(fname, @"(.)+(fblef)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefr = Regex.Match(fname, @"(.)+(fblefr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefv = Regex.Match(fname, @"(.)+(fblefv)\.tif", RegexOptions.IgnoreCase).Success;

                            if (fblef)
                            {
                                frontMatterLabels.FlagStatus = "Missing recto or verso indicator in filename";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Front cover";
                                dipsCompliant = false;
                            }
                            else if (fblefr)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Front cover";


                            }
                            else if (fblefv)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Inside front cover";
                            }

                            else
                            {
                                Console.WriteLine("ERROR: SOMETHING HAS GONE BADLY WRONG WITH ORDER & LABEL GEN... CHECK WHAT");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                frontMatterLabels.FlagStatus = errString;
                                frontMatterLabels.ObjectType = "Page";
                                frontMatterLabels.Label = derivedFilename;
                                dipsCompliant = false;


                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            frontMatterLabels.FileName = fname;
                            frontMatterLabels.FlagStatus = errString;
                            frontMatterLabels.ObjectType = "Page";
                            frontMatterLabels.Label = derivedFilename;
                            dipsCompliant = false;



                        }
                        frontMatter.Add(frontMatterLabels);
                    }
                }
                if (cFrontFlysheets.Any())
                {

                    FFSExists = true;

                        foreach (string fname in cFrontFlysheets)
                    {
                        FileLabels frontFlysheetLabels = new FileLabels();

                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;
                        // replace the string
                        string derivedFilename = split2.Last();
                        string trimmedName = derivedFilename.Trim('f', 's');
                        string noZerosName = trimmedName.TrimStart('0');
                        
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";
                            //orderNumber.Add(Int32.Parse(noZerosName));

                        

                        string matchString = derivedShelfmark + @"_(fs[0-9]+[rv])\.tif";
                        var match = Regex.Match(fname, matchString, RegexOptions.IgnoreCase);
                        var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif", @".tif");
                        var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);
                        if (match.Success || altMatch.Success)
                        { 
                           

                            frontFlysheetLabels.FileName = fname;

                            var fsr = Regex.Match(fname, @"(.)+((fs)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsv = Regex.Match(fname, @"(.)+((fs)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            
                            if (fsr)
                            {
                                frontFlysheetLabels.FlagStatus = "";
                                frontFlysheetLabels.ObjectType = "Flysheet";
                                string flysheetLabelString = "Front flysheet " + noZerosName;
                                frontFlysheetLabels.Label = flysheetLabelString;
                                int subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));

                                frontFlysheetLabels.SubOrder = subOrderNumber + (subOrderNumber -2); 
                                // based on the formula order number = n + (n-2) for recto
                                // and o = n + (n-1) for verso
                                // ensures images 1r are number zero
                                // and image 1v is number 1, etc.

                            }
                            else if (fsv)
                            {
                                frontFlysheetLabels.FlagStatus = "";
                                frontFlysheetLabels.ObjectType = "Flysheet";
                                string flysheetLabelString = "Front flysheet " + noZerosName;
                                frontFlysheetLabels.Label = flysheetLabelString;
                                int subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));
                                frontFlysheetLabels.SubOrder = subOrderNumber + (subOrderNumber - 1);

                            }
                            else
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                frontFlysheetLabels.FlagStatus = errString;
                                frontFlysheetLabels.ObjectType = "Page";
                                frontFlysheetLabels.Label = derivedFilename;
                                dipsCompliant = false;


                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            frontFlysheetLabels.FileName = fname;
                            frontFlysheetLabels.FlagStatus = errString;
                            frontFlysheetLabels.ObjectType = "Page";
                            frontFlysheetLabels.Label = derivedFilename;
                            dipsCompliant = false;



                        }
                        frontFlysheets.Add(frontFlysheetLabels);

                        }

                }
                if (cFolios.Any())
                {

                    FOLExists = true;
                    foreach (string fname in cFolios)
                    {
                        FileLabels folioLabels = new FileLabels();

                        List<String> fmat = new List<String>();
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        folderDerivedShelfmark = derivedShelfmark;

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");

                        // replace the string
                        string derivedFilename = split2.Last();
                        string trimmedName = derivedFilename.Trim('f');
                        string noZerosName = trimmedName.TrimStart('0');
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";

                        string matchString = derivedShelfmark + @"_(f)([0-9])+([rv])\.tif";
                        List<String> fols = new List<String>();
                        var match = Regex.Match(fname, matchString, RegexOptions.IgnoreCase);

                        var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif", @".tif");
                        var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);
                        if (match.Success || altMatch.Success)
                        {
                            folioLabels.FileName = fname;
                            var fr = Regex.Match(fname, @"(.)+((f)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fv = Regex.Match(fname, @"(.)+((f)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;

                            if (fr)
                            {
                                folioLabels.FlagStatus = "";
                                folioLabels.ObjectType = "Page";
                                string frString = "f. " + noZerosName;
                                folioLabels.Label = frString;
                                var subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));
                                folioLabels.SubOrder = subOrderNumber + (subOrderNumber - 2);


                            }
                            else if (fv)
                            {
                                folioLabels.FlagStatus= ""; // little bit redundant, remove after testing this works
                                folioLabels.ObjectType="Page";
                                string frString = "f. " + noZerosName;
                                folioLabels.Label=frString;
                                var subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));
                                folioLabels.SubOrder = subOrderNumber + (subOrderNumber - 1);


                            }
                            else
                            {
                                Console.WriteLine("ERROR: Folio outside of common DIPS string range. Investigate");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                folioLabels.FlagStatus=errString;
                                folioLabels.ObjectType="Page";
                                folioLabels.Label = derivedFilename;
                                dipsCompliant = false;

                            }

                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            folioLabels.FileName = fname;
                            folioLabels.FlagStatus = errString;
                            folioLabels.ObjectType = "Page";
                            folioLabels.Label=derivedFilename;
                            dipsCompliant = false;

                        }
                        folios.Add(folioLabels);
                    }
                }
                // sort out numeric foliation here
                if (cNumericFolios.Any())
                {

                    numFOLExists = true;
                    int orderCheckNum = 0;
                    foreach (string fname in cNumericFolios)
                    {
                        FileLabels numFLabels = new FileLabels();

                        List<String> nfols = new List<String>();
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        theShelfmark = derivedShelfmark;
                        string derivedFilename = split2.Last();
                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;
                        string noZerosName = derivedFilename.TrimStart('0');
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";
                        
                        if (derivedShelfmark.Contains(inputShelfmark) || inputShelfmark.Contains(derivedShelfmark))
                        {
                            int fileNameNumber = Int32.Parse(noZerosName);
                            numFLabels.SubOrder = fileNameNumber;
                            if (fileNameNumber - orderCheckNum != 1)
                            {
                                orderCheck = false;
                            }
                            orderCheckNum = fileNameNumber;

                            string matchString = theShelfmark + @"_([0-9])+\.tif";
                            var match = Regex.Match(fname, matchString, RegexOptions.IgnoreCase);

                            var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif", @".tif");
                            var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);


                            numFLabels.FileName = fname;
                            numFLabels.FlagStatus = "";
                            numFLabels.ObjectType = "Page";
                            numFLabels.Label = derivedFilename;

                            if (!match.Success && !altMatch.Success)
                            {

                                Console.WriteLine("ERROR: Doesn't match numeric filenaming pattern");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                numFLabels.FlagStatus = errString;
                                dipsCompliant = false;


                            }
                            else if (!orderCheck)
                            {
                                string errString = "Non-consecutive numbering in image name";
                                numFLabels.FlagStatus = errString;
                            }

                        }
                        else
                        {
                            Console.WriteLine("ERROR: Doesn't match shelfmark filenaming pattern");
                            numFLabels.FileName = fname;
                            numFLabels.FlagStatus = "File name not formed from shelfmark";
                            numFLabels.ObjectType = "Unknown";
                            numFLabels.Label = derivedFilename;
                            dipsCompliant = false;

                        }

                        numericFiles.Add(numFLabels);
                    }
                }

                if (cEndFlysheets.Any())
                {

                    EFSExists = true;
                    foreach (string fname in cEndFlysheets)
                    {
                        FileLabels efsLabels = new FileLabels();

                        List<String> fmat = new List<String>();
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;

                        // replace the string
                        string derivedFilename = split2.Last();
                        string trimmedName = derivedFilename.Trim('f','s','e');
                        string noZerosName = trimmedName.TrimStart('0');
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";

                        List<String> efs = new List<String>();
                        string matchString = derivedShelfmark + @"_((fse)[0-9]+[rv])\.tif";

                        var match = Regex.Match(fname, matchString, RegexOptions.IgnoreCase);

                        var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif", @".tif");
                        var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);
                        if (match.Success || altMatch.Success)
                        {
                            efsLabels.FileName = fname;
                            var fser = Regex.Match(fname, @"(.)+((fse)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsev = Regex.Match(fname, @"(.)+((fse)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            if (fser)
                            {
                                efsLabels.FlagStatus = ""; // error string
                                efsLabels.ObjectType = "Flysheet";
                                string frString = "Back flysheet " + noZerosName;
                                efsLabels.Label = frString;
                                var subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));
                                efsLabels.SubOrder = subOrderNumber + (subOrderNumber - 2);

                            }
                            else if (fsev)
                            {
                                efsLabels.FlagStatus = ""; // error string
                                efsLabels.ObjectType = "Flysheet";
                                string frString = "Back flysheet " + noZerosName;
                                efsLabels.Label = frString;
                                var subOrderNumber = Int32.Parse(noZerosName.TrimEnd('r', 'v'));
                                efsLabels.SubOrder = subOrderNumber + (subOrderNumber - 1);
                            }
                            else
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                efsLabels.FlagStatus = errString; // error string
                                efsLabels.ObjectType = "Flysheet";
                                efsLabels.Label = derivedFilename;
                                dipsCompliant = false;

                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            efsLabels.FileName = fname;
                            efsLabels.FlagStatus = errString;
                            efsLabels.ObjectType = "Page";
                            efsLabels.Label = derivedFilename;
                            dipsCompliant = false;


                        }
                        endFlysheets.Add(efsLabels);
                    }
                }
                if (cEndMatter.Any())
                {

                    EMExists = true;
                    foreach (string fname in cEndMatter)
                    {
                        FileLabels emLabels = new FileLabels();

                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        string derivedFilename = split2.Last();

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;

                        string matchString = derivedShelfmark + @"_(((fb)((rig)|(rigv)|(rigr)|(spi))))\.tif";

                        List<String> ema = new List<String>();
                        var match = Regex.Match(fname,matchString, RegexOptions.IgnoreCase);

                        var cleanedFName = fname.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~").Replace("_tif", @".tif");
                        var altMatch = Regex.Match(cleanedFName, matchString, RegexOptions.IgnoreCase);
                        if (match.Success || altMatch.Success)
                        {
                            emLabels.FileName = fname;
                            var fbrig = Regex.Match(fname, @"(.)+(fbrig)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbrigr = Regex.Match(fname, @"(.)+(fbrigr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbrigv = Regex.Match(fname, @"(.)+(fbrigv)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbspi = Regex.Match(fname, @"(.)+(fbspi)\.tif", RegexOptions.IgnoreCase).Success;
                            if (fbrig)
                            {
                                emLabels.FlagStatus = "Missing recto or verso indicator in filename";
                                emLabels.ObjectType = "Cover";
                                emLabels.Label = "Back cover";
                                dipsCompliant = false;
                            }
                            else if (fbrigr)
                            {
                                emLabels.FlagStatus = "";
                                emLabels.ObjectType = "Cover";
                                emLabels.Label = "Inside back cover";
                            }
                            else if (fbrigv)
                            {
                                emLabels.FlagStatus = "";
                                emLabels.ObjectType = "Cover";
                                emLabels.Label = "Back cover";
                            }
                            else if (fbspi)
                            {
                                emLabels.FlagStatus = "";
                                emLabels.ObjectType = "Cover";
                                emLabels.Label = "Spine";
                            }
                            else // no match for any of these 'usual' cases
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                emLabels.FlagStatus = errString;
                                emLabels.ObjectType = "Page";
                                emLabels.Label = derivedFilename;
                                dipsCompliant = false;

                            }


                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            emLabels.FileName = fname;
                            emLabels.FlagStatus = errString;
                            emLabels.ObjectType = "Page";
                            emLabels.Label = derivedFilename;
                            dipsCompliant = false;

                        }
                        endMatter.Add(emLabels);
                    }
                }

                if (otherFiles.Any()) // checks for control shot etc and anything out of the ordinary
                {

                    foreach (string fname in otherFiles)
                    {
                        FileLabels unclassifiedLabels = new FileLabels();

                        List<String> others = new List<String>();
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");

                        theShelfmark = derivedShelfmark;
                        string derivedFilename = split2.Last();

                        folderDerivedShelfmark = derivedShelfmark;
                        // check derived shelfmark against inputShelfmark
                        var checkShelfmark = inputShelfmark;

                        

                        if (!theShelfmark.Contains(inputShelfmark) && !inputShelfmark.Contains(theShelfmark))
                        {
                            unclassifiedLabels.FileName = fname;
                            unclassifiedLabels.FlagStatus = "Does not meet DIPS standards and filename not derived from shelfmark";
                            unclassifiedLabels.ObjectType = "Unknown";
                            unclassifiedLabels.Label = derivedFilename;
                            dipsCompliant = false;

                        }
                        else if (derivedFilename.ToUpper().ToLower().Contains("control"))
                        {
                            unclassifiedLabels.FileName = fname;
                            unclassifiedLabels.FlagStatus = "Does not meet DIPS standards - possible control shot";
                            unclassifiedLabels.ObjectType = "Control shot";
                            unclassifiedLabels.Label = derivedFilename;
                            dipsCompliant = false;
                        }
                        else
                        {
                            unclassifiedLabels.FileName = fname;
                            unclassifiedLabels.FlagStatus = "Does not meet DIPS standards";
                            unclassifiedLabels.ObjectType = "Unknown";
                            unclassifiedLabels.Label = derivedFilename;
                            dipsCompliant = false;
                        }

                        unclassified.Add(unclassifiedLabels);
                    }
                }

                // Sort everything by filename at this point
                frontMatter = frontMatter.OrderBy(o => o.FileName).ToList();
                frontFlysheets = frontFlysheets.OrderBy(o => o.SubOrder).ToList(); // flysheets need to be ordered numerically
                folios = folios.OrderBy(o => o.SubOrder).ToList();
                endFlysheets = endFlysheets.OrderBy(o => o.SubOrder).ToList();
                endMatter = endMatter.OrderBy(o => o.FileName).ToList();
                numericFiles = numericFiles.OrderBy(o => o.SubOrder).ToList(); // was filename but can sort numeric files by SubOrder now!
                unclassified = unclassified.OrderBy(o => o.FileName).ToList();

                // Flagging DIPS compliance mismatches:
                bool containsDIPSNames = false;
                if (FMExists || FFSExists || FOLExists || EFSExists || EMExists)
                {
                    containsDIPSNames = true;
                }

                if (containsDIPSNames && numFOLExists) // if numerically labelled folios exist alongside any DIPS compliant names...
                {
                    
                    Console.WriteLine("Mixture of DIPS-compliant and numerical filenames in shelfmark {0}", folderDerivedShelfmark);
                    String columnName = "DIPS_Compliance";
                    String message = "Mixture";
                    if (!dipsCompliant)
                    {
                        message += "; Contains DIPS-invalid filenames";
                    }
                    if (!orderCheck)
                    {
                        message += "; Non-consecutive filenames";
                    }
                    try
                    {
                        SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                        SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing file-level DIPS compliance status to sharepoint for shelfmark {0}. \nException {1}", theShelfmark, ex);
                    }
                }
                else if (!dipsCompliant)
                {
                    Console.WriteLine("{0} contains filenames that do not meet DIPS standards", folderDerivedShelfmark);
                    String columnName = "DIPS_Compliance";
                    String message = "Invalid filenames";
                    try
                    {
                        SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                        SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing file-level DIPS compliance status to sharepoint for shelfmark {0}. \nException {1}", theShelfmark, ex);
                    }
                }
                else // Things are good!
                {
                    String columnName = "DIPS_Compliance";
                    String message = "Full";
                    try
                    {
                        SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                        SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing file-level DIPS compliance status to sharepoint for shelfmark {0}. \nException {1}", theShelfmark, ex);
                    }

                }

                foreach (FileLabels fm in frontMatter)
                {
                    allFilesSorted.Add(fm);
                }
                foreach (FileLabels ffs in frontFlysheets)
                {
                    allFilesSorted.Add(ffs);
                }
                foreach (FileLabels fol in folios)
                {
                    allFilesSorted.Add(fol);
                }
                // Add in the numerically labelled files if they exist, we've sorted out the error flags here anyway...
                foreach (FileLabels nfol in numericFiles)
                {
                    allFilesSorted.Add(nfol);
                }

                foreach (FileLabels efs in endFlysheets)
                {
                    allFilesSorted.Add(efs);
                }
                foreach (FileLabels em in endMatter)
                {
                    allFilesSorted.Add(em);
                }
                foreach (FileLabels oth in unclassified)
                {
                    allFilesSorted.Add(oth);
                }

            } // if at least some DIPS-compliant filenames exist
            else // is fully non-DIPS compliant and just has numerical filenames, so just sort this normally
            {
                bool compliantFilenames = true;
                FileLabels numFile = new FileLabels();
                List<String> sortedFilenames = fileNames.OrderBy(x => x).Select(x => x.ToString()).ToList();
                int orderCheckNum = 0;
                foreach (var sfn in sortedFilenames)
                {
                    List<String> nums = new List<String>();
                   
                    string[] split = sfn.Split('.');
                    string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                    string fileExtension = split.Last(); // tif
                    string[] split2 = shelfmark_filename.Split('_');
                    string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                    string derivedFilename = split2.Last();
                    string noZerosName = derivedFilename.TrimStart('0');
                    noZerosName = noZerosName.Length > 0 ? noZerosName : "0";
                    int fileNameNumber = Int32.Parse(noZerosName);
                    if (fileNameNumber - orderCheckNum != 1)
                    {
                        orderCheck = false;
                        compliantFilenames = false;
                    }
                    orderCheckNum = fileNameNumber;

                    string matchString = theShelfmark + @"_([0-9])+\.tif";

                    var match = Regex.Match(sfn, matchString, RegexOptions.IgnoreCase);
                    
                    numFile.FileName = sfn;
                    numFile.FlagStatus = ""; // errorString
                    numFile.ObjectType = "Page";
                    numFile.Label = noZerosName; // just get the number from the filename

                    if (!match.Success)
                    {
                        compliantFilenames = false;
                        Console.WriteLine("ERROR: Doesn't match numeric filenaming pattern");
                        string errString = "Unexpected characters in filename. Flag for investigation";
                        numFile.FlagStatus = errString;

                    }

                    allFilesSorted.Add(numFile);
                }

                if (compliantFilenames)
                {

                    String columnName = "DIPS_Compliance";
                    String message = "Full (numeric)";
                    try
                    {
                        SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                        SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing file-level DIPS compliance status to sharepoint for shelfmark {0}. \nException {1}", theShelfmark, ex);
                    }

                }
                else
                {
                    String columnName = "DIPS_Compliance";
                    String message = "Invalid (numeric)";
                    if (!orderCheck)
                    {
                        message += "; Non-consecutive filenames";
                    }
                    try
                    {
                        SharepointTools.CreateSharepointColumn(spURL, "Digitisation Workflow", columnName);
                        SharepointTools.WriteToSharepointColumnByID(spURL, spList, columnName, theShelfmark, itemID, message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Problem writing file-level DIPS compliance status to sharepoint for shelfmark {0}. \nException {1}", theShelfmark, ex);
                    }
                }

            }

            // At this stage you have allFilesSorted as a list-of-lists with
           // filename , flagStatus, objectType, Label 
           //- flagStatus is a string that is either empty (all good!) or contains an error message
           // objectType is jut page, cover, flysheet etc
           // label is "Inside back cover", "folio 5v" etc
           for(int i = 0; i<allFilesSorted.Count; i++)
            {
                string orderNumber = (i + 1).ToString();
                allFilesSorted[i].OrderNumber = orderNumber;
            }

           

           // now (each item in) allFilesSorted contains a FileLabels object with an associated
           //filename
           //flagStatus
           //objectType
           //label
           //order number
           // (all strings)
            return allFilesSorted;
        }

       

        private static bool writeFileLabelsToCSV(List<FileLabels> ShelfmarkFilesLabels, String outFolder)
        {
            bool fError = false;

            try // to write the csv...
            {
                List<String> strHeaders = new List<string>{"File","Order","Type","Label"};
                System.Text.UnicodeEncoding uce = new System.Text.UnicodeEncoding();
                string fNameString = "ImageOrder";
                string outPath = outFolder + @"\"+fNameString+".csv";

                if (File.Exists(outPath))
                {
                    String lastModified = File.GetLastWriteTime(outPath).ToString("yyyyMMdd_HH-mm-ss");
                    string oldFilePath = outFolder + @"\" + fNameString + "_" + lastModified + ".csv";

                    try
                    {
                        File.Move(outPath, oldFilePath); // moves existing imageorder.csv file to include the last modified time
                                                         // leaves newest created file as 'imageorder.csv'
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Could not move existing ImageOrder.csv from {0} to {1}.\nException: {2}", outPath, oldFilePath, ex);
                    }

                }
                    

                using (var sr = new StreamWriter(outPath, false, uce))
                {
                    using (var csvFile = new CsvHelper.CsvWriter(sr, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        csvFile.Configuration.Delimiter = "\t";
                        //csvFile.Configuration.HasExcelSeparator = true;

                        foreach (var header in strHeaders)
                        {
                            csvFile.WriteField(header);
                        }
                        csvFile.NextRecord(); // skips to next line...
                        foreach (var record in ShelfmarkFilesLabels)
                        { 
                            csvFile.WriteField(record.FileName); // filename
                            csvFile.WriteField(record.OrderNumber); // order number
                            csvFile.WriteField(record.ObjectType); // object type
                            csvFile.WriteField(record.Label); // label
                            csvFile.WriteField(record.FlagStatus); // error flag status
                        
                            if (ShelfmarkFilesLabels.IndexOf(record) != ShelfmarkFilesLabels.Count - 1)
                            {
                                csvFile.NextRecord();
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

        public static String GetXMLVersionNumber(String fileName)
        {
            XDocument doc = XDocument.Load(fileName);
            string version = doc.Declaration.Version;
            return version;

        }
    }
}


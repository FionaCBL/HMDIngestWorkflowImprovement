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


        public string Filename { get; set; }
        public string FlagStatus { get; set; }
        public string ObjectType { get; set; }
        public string Label { get; set; }
        public string OrderNumber { get; set; }


        public FileLabels(string filename, string flagstatus, string objtype, string label, string ordernum)
        {
            Filename = filename;
            FlagStatus = flagstatus;
            ObjectType = objtype;
            Label = label;
            OrderNumber = ordernum;

        }
        public FileLabels()
        {
            Filename = null;
            FlagStatus = null;
            ObjectType = null;
            Label = null;
            OrderNumber = null;
        }

    }
    class InputOrderSpreadsheetTools
    {
        public static List<List<FileLabels>> listAllShelfmarkFilesTIFXML(List<HMDObject> sharepointOut, String env, String spURL, String spList)
        {
            List<List<String>> sourceFolderXMLs = new List<List<String>>(); // maybe don't need?
            List<List<FileLabels>> allShelfmarkTIFAndLabels = new List<List<FileLabels>>();


            foreach(var item in sharepointOut)
            {
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

                    Console.WriteLine("Source Folder: {0}", sourceFolder);
                    try
                    {
                        if (sourceFolder.ToUpper().ToLower().Contains("tif"))
                        {
                            tifFolder = sourceFolder;
                        }
                        else
                        {

                            var subFolders = Directory.GetDirectories(sourceFolder);
                            foreach (var subFolder in subFolders)
                            {
                                Console.WriteLine("Testing subFolder: {0}", subFolder);
                                if (subFolder.ToUpper().ToLower().Contains("tif"))
                                {
                                    tifFolder = subFolder;
                                    Console.WriteLine("Found subfolder for folder {0}", sourceFolder);
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
                        shelfmarkLabels = mapFileNameToLabels(shelfmark,Files, tifFolder);
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
                                folderShelfmark = folderShelfmark.TrimEnd('_');


                                string testTifFolder = tifFolder.Split(new string[] { folderShelfmark }, 2, StringSplitOptions.None)[1];

                                outFolder += testTifFolder;
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
                            string outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            outFolder += @"\HMDSharepoint_ImgOrderCSVs";
                            string SM_folderFormat = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");

                            string folderShelfmark = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                            try
                            {
                                folderShelfmark = folderShelfmark.TrimEnd('_'); // just testing this, might remove in future.

                                string testTifFolder = tifFolder.Split(new string[] { folderShelfmark }, 2, StringSplitOptions.None)[1];

                                outFolder += testTifFolder;
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
                                outFolder += @"\HMDSharepoint_ImgOrderCSVs" + @"\"+SM_folderFormat;
                                if (!Directory.Exists(outFolder))
                                {
                                    Directory.CreateDirectory(outFolder);
                                }                        // Now write this to a CSV

                                Assert.IsTrue(writeFileLabelsToCSV(shelfmarkLabels, outFolder));

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
                                    //Assert.IsTrue(SharepointTools.WriteToSharepointColumnBySingleShelfmark(spURL, spList, "ALTOXMLCheck", shelfmark, xmlErrMessage));
                                }
                                else if (env == "prod")
                                {
                                    Console.WriteLine("Holding off on populting columns in Sharepoint prod version for now");
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

                        FileLabels errorList = new FileLabels(shelfmark, null, null, null, null );
                        errorList.FlagStatus = "TIF folder not found";

                        continue; // use continue for now, but will need to write out invalid path to a variable at some point
                    }
                } // is source folder valid? 
                else // source folder was never valid
                {
                    FileLabels errorList = new FileLabels(shelfmark, null, null, null, null);
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

        private static List<FileLabels> mapFileNameToLabels(String inputShelfmark, FileInfo[] Files, String tifFolders)
        {

            // Order labels will take a couple of sweeps - one to get front and back matter and then another to do a fine sort of the front and back matter
            List<String> shelfmarkLabels = new List<String>();
            string theShelfmark = "";
            inputShelfmark = inputShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");


            List<string> fileNames = Files.Select(x => x.Name).ToList();

            // Define regular expressions to search for
            // 'initial' versions perform a looser search
            string iFrontMatterReString = @"(.)+(((fble)((fv)|(fr)))|((fs)[0-9]+(.)+))\.tif";
            var initialFrontMatterRegex = new Regex(iFrontMatterReString, RegexOptions.IgnoreCase);

            string iFolioReString = @"(.)+(f)([0-9])+(.)+\.tif";
            var initialFolioRegex = new Regex(iFolioReString, RegexOptions.IgnoreCase);

            string iEndFlysheetsReString = @"(.)+((fse)[0-9]+(.)+)\.tif";
            var initialEndFlysheetsRegex = new Regex(iEndFlysheetsReString, RegexOptions.IgnoreCase);


            string iEndMatterReString = @"(.)+(((fb)((rigv)|(rigr)|(spi))))\.tif";
            var initialEndMatterRegex = new Regex(iEndMatterReString, RegexOptions.IgnoreCase);

            string iNumericFolioReString = @"(.)+_([0-9])+\.tif";
            var initialNumericFolioRegex = new Regex(iNumericFolioReString, RegexOptions.IgnoreCase);

            // Sort into front matter, end flysheets, end matter and folios
            // Control shots jut shouldn't get picked up at all by any of these
            // Do need to pick up numerically labelled files though


            // Candidates for each section:
            List<string> cFrontMatter = fileNames.Where(f => initialFrontMatterRegex.IsMatch(f)).ToList();
            List<string> cEndFlysheets = fileNames.Where(f => initialEndFlysheetsRegex.IsMatch(f)).ToList();
            List<string> cEndMatter = fileNames.Where(f => initialEndMatterRegex.IsMatch(f)).ToList();
            List<string> cFolios = fileNames.Where(f => initialFolioRegex.IsMatch(f)).ToList();
            List<string> cNumericFolios = fileNames.Where(f => initialNumericFolioRegex.IsMatch(f)).ToList();

            List<string> foundItems = cFrontMatter.Concat(cFolios).Concat(cNumericFolios).Concat(cEndFlysheets).Concat(cEndMatter).ToList();


            List<String> otherFiles = (from e in (fileNames.Concat(foundItems))
                                       where !foundItems.Contains(e) select e).ToList(); // get everything not found (should be control shots etc)
                                                                                         // is the above worth doing?


            List<FileLabels> allFilesSorted = new List<FileLabels>(); // this is what you're returning later

            List<FileLabels> frontMatter = new List<FileLabels>();
            List<FileLabels> endFlysheets = new List<FileLabels>();
            List<FileLabels> endMatter = new List<FileLabels>();
            List<FileLabels> folios = new List<FileLabels>();
            List<FileLabels> numericFiles = new List<FileLabels>();

            string folderDerivedShelfmark = "";
        
            if (cFrontMatter.Any() | cFolios.Any() | cEndMatter.Any() | cEndFlysheets.Any())
            {
                // you can be pretty sure its DIPs compliant if you see any titles or any numbered folios

                bool FMExists = false;
                bool FOLExists = false;
                bool EFSExists = false;
                bool EMExists = false;
                bool numFOLExists = false;
                if (cFrontMatter.Any())
                {
                    FileLabels frontMatterLabels = new FileLabels();

                    FMExists = true;
                    foreach (string fname in cFrontMatter)
                    {
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



                        string matchString = derivedShelfmark+@"_(((fble)((fv)|(fr)))|((fs)[0-9]+[rv]))\.tif";
                        var match = Regex.Match(fname,matchString, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            frontMatterLabels.Filename = fname;
                            var fblef = Regex.Match(fname, @"(.)+(fblef)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefr = Regex.Match(fname, @"(.)+(fblefr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefv = Regex.Match(fname, @"(.)+(fblefv)\.tif", RegexOptions.IgnoreCase).Success;
                            var fsr = Regex.Match(fname, @"(.)+((fs)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsv = Regex.Match(fname, @"(.)+((fs)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            if (fblef)
                            {
                                frontMatterLabels.FlagStatus = "Missing recto or verso indicator in filename";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Front cover";
                            }
                            if (fblefr)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Front cover";

                             
                            }
                            else if (fblefv)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Cover";
                                frontMatterLabels.Label = "Front cover inside";
                            }
                            else if (fsr)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Flysheet"; 
                                string flysheetLabelString = "Front flysheet " + noZerosName;
                                frontMatterLabels.Label = flysheetLabelString;
                            }
                            else if (fsv)
                            {
                                frontMatterLabels.FlagStatus = "";
                                frontMatterLabels.ObjectType = "Flysheet";
                                string flysheetLabelString = "front flysheet " + noZerosName;
                                frontMatterLabels.Label = flysheetLabelString;
                            }
                            else
                            {
                                Console.WriteLine("ERROR: SOMETHING HAS GONE BADLY WRONG WITH ORDER & LABEL GEN... CHECK WHAT");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                frontMatterLabels.FlagStatus = errString;
                                frontMatterLabels.ObjectType = "Page";
                                frontMatterLabels.Label = derivedFilename;

                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            frontMatterLabels.Filename = fname;
                            frontMatterLabels.FlagStatus = errString;
                            frontMatterLabels.ObjectType = "Page";
                            frontMatterLabels.Label = derivedFilename;
                           

                        }
                        frontMatter.Add(frontMatterLabels);
                    }
                }
                if (cFolios.Any())
                {
                    FileLabels folioLabels = new FileLabels();

                    FOLExists = true;
                    foreach (string fname in cFolios)
                    {
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
                        if (match.Success)
                        {
                            folioLabels.Filename = fname;
                            var fr = Regex.Match(fname, @"(.)+((f)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fv = Regex.Match(fname, @"(.)+((f)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;

                            if (fr)
                            {
                                folioLabels.FlagStatus = "";
                                folioLabels.ObjectType = "Page";
                                string frString = "Folio " + noZerosName;
                                folioLabels.Label = frString;

                            }
                            else if (fv)
                            {
                                folioLabels.FlagStatus= ""; // little bit redundant, remove after testing this works
                                folioLabels.ObjectType="Page";
                                string frString = "Folio " + noZerosName;
                                folioLabels.Label=frString;
                            }
                            else
                            {
                                Console.WriteLine("ERROR: Folio outside of common DIPS string range. Investigate");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                folioLabels.FlagStatus=errString;
                                folioLabels.ObjectType="Page";
                                folioLabels.Label = derivedFilename;
                            }

                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            folioLabels.Filename = fname;
                            folioLabels.FlagStatus = errString;
                            folioLabels.ObjectType = "Page";
                            folioLabels.Label=derivedFilename;
                        }
                        folios.Add(folioLabels);
                    }
                }
                // sort out numeric foliation here
                if (cNumericFolios.Any())
                {
                    FileLabels numFLabels = new FileLabels();

                    numFOLExists = true;
                    foreach (string fname in cNumericFolios)
                    {
                        List<String> nfols = new List<String>();
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        theShelfmark = derivedShelfmark;
                        string derivedFilename = split2.Last();
                        folderDerivedShelfmark = derivedShelfmark;


                        string matchString = theShelfmark+@"_([0-9])+\.tif";

                        var match = Regex.Match(fname, matchString, RegexOptions.IgnoreCase);

                        numFLabels.Filename = fname;
                        numFLabels.FlagStatus = "";
                        numFLabels.ObjectType = "Page";
                        numFLabels.Label = derivedFilename;

                        if (!match.Success)
                        {
                       
                            Console.WriteLine("ERROR: Doesn't match numeric filenaming pattern");
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            numFLabels.FlagStatus = errString;
 
                        }
                        numericFiles.Add(numFLabels);
                    }
                }

                if (cEndFlysheets.Any())
                {
                    FileLabels efsLabels = new FileLabels();

                    EFSExists = true;
                    foreach (string fname in cEndFlysheets)
                    {
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
                        if (match.Success)
                        {
                            efsLabels.Filename = fname;
                            var fser = Regex.Match(fname, @"(.)+((fse)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsev = Regex.Match(fname, @"(.)+((fse)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            if (fser)
                            {
                                efsLabels.FlagStatus = ""; // error string
                                efsLabels.ObjectType = "Flysheet";
                                string frString = "Back flysheet " + noZerosName;
                                efsLabels.Label = frString;
                            }
                            else if (fsev)
                            {
                                efsLabels.FlagStatus = ""; // error string
                                efsLabels.ObjectType = "Flysheet";
                                string frString = "Back flysheet " + noZerosName;
                                efsLabels.Label = frString;
                            }
                            else
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                efsLabels.FlagStatus = errString; // error string
                                efsLabels.ObjectType = "Flysheet";
                                efsLabels.Label = derivedFilename;
                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            efsLabels.Filename = fname;
                            efsLabels.FlagStatus = errString;
                            efsLabels.ObjectType = "Page";
                            efsLabels.Label = derivedFilename;

                        }
                        endFlysheets.Add(efsLabels);
                    }
                }
                if (cEndMatter.Any())
                {
                    FileLabels emLabels = new FileLabels();

                    EMExists = true;
                    foreach (string fname in cEndMatter)
                    {
                        string[] split = fname.Split('.');
                        string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                        string fileExtension = split.Last(); // tif
                        string[] split2 = shelfmark_filename.Split('_');
                        string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                        string derivedFilename = split2.Last();

                        derivedShelfmark = derivedShelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                        folderDerivedShelfmark = derivedShelfmark;

                        string matchString = derivedShelfmark + @"_(((fb)((rigv)|(rigr)|(spi))))\.tif";

                        List<String> ema = new List<String>();
                        var match = Regex.Match(fname,matchString, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            emLabels.Filename = fname;
                            var fbrigr = Regex.Match(fname, @"(.)+(fbrigr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbrigv = Regex.Match(fname, @"(.)+(fbrigv)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbspi = Regex.Match(fname, @"(.)+(fbspi)\.tif", RegexOptions.IgnoreCase).Success;

                            if (fbrigr)
                            {
                                emLabels.FlagStatus = "";
                                emLabels.ObjectType = "Cover";
                                emLabels.Label = "Back cover inside";
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
                            }


                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            emLabels.Filename = fname;
                            emLabels.FlagStatus = errString;
                            emLabels.ObjectType = "Page";
                            emLabels.Label = derivedFilename;
                        }
                        endMatter.Add(emLabels);
                    }
                }

                // Sort everything by filename at this point
                frontMatter = frontMatter.OrderBy(o => o.Filename).ToList();
                folios = folios.OrderBy(o => o.Filename).ToList();
                endFlysheets = endFlysheets.OrderBy(o => o.Filename).ToList();
                endMatter = endMatter.OrderBy(o => o.Filename).ToList();
                numericFiles = numericFiles.OrderBy(o => o.Filename).ToList();


                // Flagging DIPS compliance mismatches:
                bool containsDIPSNames = false;
                if (FMExists || FOLExists || EFSExists || EMExists)
                {
                    containsDIPSNames = true;
                }

                if (containsDIPSNames && numFOLExists) // if numerically labelled folios exist alongside any DIPS compliant names...
                {
                    //TODO:
                    // Need this to trigger some writing to sharepoint - Not yet working.
                    // This does actually work now, put it in.

                    Console.WriteLine("Mixture of DIPS-compliant and non-compliant filenames in shelfmark {0}", folderDerivedShelfmark);
                    // Just write this to console for now! Also would write to sharepoint for this shelfmark

                }

                foreach (FileLabels fm in frontMatter)
                {
                    allFilesSorted.Add(fm);
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

            } // if at least some DIPS-compliant filenames exist
            else // is fully non-DIPS compliant and just has numerical filenames, so just sort this normally
            {
                FileLabels numFile = new FileLabels();

                List<String> sortedFilenames = fileNames.OrderBy(x => x).Select(x => x.ToString()).ToList();
                foreach (var sfn in sortedFilenames)
                {
                    List<String> nums = new List<String>();
                    numFile.Filename = sfn;
                    numFile.FlagStatus = ""; // errorString
                    numFile.ObjectType = "Page";
                    string[] split = sfn.Split('.');
                    string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                    string fileExtension = split.Last(); // tif
                    string[] split2 = shelfmark_filename.Split('_');
                    string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                    string derivedFilename = split2.Last();
                    string noZerosName = derivedFilename.TrimStart('0');
                    noZerosName = noZerosName.Length > 0 ? noZerosName : "0";
                    numFile.Label = noZerosName; // just get the number from the filename

                    allFilesSorted.Add(numFile);
                }
            }

            // At this stage you have allFilesSorted as a list-of-lists with
           // filename , flagStatus, objectType, Label 
           //- flagStatus is a string that is either empty (all good!) or contains an error message
           // objectType is jut page, cover, flysheet etc
           // label is "back cover inside", "folio 5v" etc
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
                const char sep = '\t';
                List<String> strHeaders = new List<string>{"File","Order","Type","Label"};
                System.Text.UnicodeEncoding uce = new System.Text.UnicodeEncoding();
                string fNameString = "ImageOrder";
                string outPath = outFolder + @"\"+fNameString+".csv";

                if (File.Exists(outPath))
                {
                    var time = DateTime.Now;
                    string formattedTime = time.ToString("yyyyMMdd_HH-mm-ss");
                    string altOutPath = outFolder + @"\" + fNameString + "_" + formattedTime + ".csv";
                    outPath = altOutPath;

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
                            csvFile.WriteField(record.Filename); // filename
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


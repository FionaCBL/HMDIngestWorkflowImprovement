using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    class InputOrderSpreadsheetTools
    {
        public static List<List<List<String>>> listAllShelfmarkFilesTIFXML(List<List<String>> sharepointOut, String env, String spURL, String spList)
        {
            List<List<String>> sourceFolderXMLs = new List<List<String>>(); // maybe don't need?
            List<List<List<String>>> allShelfmarkTIFAndLabels = new List<List<List<String>>>();


            for (int i = 1; i < sharepointOut.Count; i++) // need this to skip the first row (titles)
            {
                List<String> item = sharepointOut[i];
                List<String> shelfmarkTIFs = new List<String>();

                List<List<String>> shelfmarkLabels = new List<List<String>>();
                bool validPath = false;
                var itemID = item[0];
                var shelfmark = item[1];

                if (item[5] != "false")
                {
                    string sourceFolder = "";

                    if (string.IsNullOrEmpty(item[6]))
                    {
                        sourceFolder = item[2];
                    }
                    else
                    {
                        sourceFolder = item[6];
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
                        // shelfmarkLabels is a list of lists
                        // each sub-list is  for a particular file and contains:
                        //[0]: filename
                        //[1]: flagStatus
                        //[2]: objectType
                        //[3]: label
                        //[4]: order number


                         if (env == "test") // get this going for prod by sticking it in the actual tifFolder
                        {
                            string outFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            outFolder += @"\HMDSharepoint_ImgOrderCSVs\";

                            string SM_folderFormat = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");

                            string folderShelfmark = shelfmark.ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                            try
                            {
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
                                    Assert.IsTrue(SharepointTools.WriteToSharepointColumnBySingleShelfmark(spURL, spList, "ALTOXMLCheck", shelfmark, xmlErrMessage));
                                }
                                else if (env == "prod")
                                {
                                    Console.WriteLine("Holding off on populting columns in Sharepoint prod version for now");
                                }
                            }



                        }


                    }// if validPath == true


                    else // so not a valid path!
                    {
                        // need to build up a list and then add it to shelfmarkLabels
                        var errorList = new List<string> { shelfmark, null, null, null, null };
                        shelfmarkLabels.Add(errorList);

                        continue; // use continue for now, but will need to write out invalid path to a variable at some point
                    }
                }
                else
                {
                    var errorList = new List<string> { shelfmark, null, null, null, null };
                    shelfmarkLabels.Add(errorList);

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

            // returns you a List<List<List<String>>>
            // Shelfmark labels are outputted as a list of list of strings - for each shelfmark you will have 
            // a list for each file: filename, image label, order label etc
            // so shelfmark labels are a list of list of strings
            // For all shelfmarks this is then List<List<List<String>>>
        }

        private static List<List<string>> mapFileNameToLabels(String inputShelfmark, FileInfo[] Files, String tifFolders)
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

            List<List<String>> allFilesSorted = new List<List<String>>(); // make this into a class...
            //LabelledFile allFilesSorted = new LabelledFile();


            List<List<string>> frontMatter = new List<List<String>>();
            List<List<string>> endFlysheets = new List<List<String>>();
            List<List<string>> endMatter = new List<List<String>>();
            List<List<string>> folios = new List<List<String>>();
            List<List<string>> numericFiles = new List<List<String>>();

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
                    FMExists = true;
                    foreach (string fname in cFrontMatter)
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
                        string trimmedName = derivedFilename.Trim('f', 's');
                        string noZerosName = trimmedName.TrimStart('0');
                        noZerosName = noZerosName.Length > 0 ? noZerosName : "0";



                        string matchString = derivedShelfmark+@"_(((fble)((fv)|(fr)))|((fs)[0-9]+[rv]))\.tif";
                        var match = Regex.Match(fname,matchString, RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            fmat.Add(fname);
                            var fblef = Regex.Match(fname, @"(.)+(fblef)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefr = Regex.Match(fname, @"(.)+(fblefr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fblefv = Regex.Match(fname, @"(.)+(fblefv)\.tif", RegexOptions.IgnoreCase).Success;
                            var fsr = Regex.Match(fname, @"(.)+((fs)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsv = Regex.Match(fname, @"(.)+((fs)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            if (fblef)
                            {
                                fmat.Add("Missing recto or verso indicator in filename");
                                fmat.Add("Cover");
                                fmat.Add("Front cover");
                            }
                            if (fblefr)
                            {
                                fmat.Add("");
                                fmat.Add("Cover");
                                fmat.Add("Front cover");
                            }
                            else if (fblefv)
                            {
                                fmat.Add("");
                                fmat.Add("Cover");
                                fmat.Add("Front cover inside");
                            }
                            else if (fsr)
                            {
                                fmat.Add("");
                                fmat.Add("Flysheet"); 
                                string flysheetLabelString = "Front flysheet " + noZerosName;
                                fmat.Add(flysheetLabelString);
                            }
                            else if (fsv)
                            {
                                fmat.Add("");
                                fmat.Add("Flysheet");
                                string flysheetLabelString = "front flysheet " + noZerosName;
                                fmat.Add(flysheetLabelString);
                            }
                            else
                            {
                                Console.WriteLine("ERROR: SOMETHING HAS GONE BADLY WRONG WITH ORDER & LABEL GEN... CHECK WHAT");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                fmat.Add(errString);
                                fmat.Add("Page");
                                fmat.Add(derivedFilename);

                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            fmat.Add(fname);
                            fmat.Add(errString);
                            fmat.Add("Page");
                            fmat.Add(derivedFilename);

                        }
                        frontMatter.Add(fmat);
                    }
                }
                if (cFolios.Any())
                {
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

                            fols.Add(fname);
                            var fr = Regex.Match(fname, @"(.)+((f)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fv = Regex.Match(fname, @"(.)+((f)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;

                            if (fr)
                            {
                                fols.Add("");
                                fols.Add("Page");
                                string frString = "Folio " + noZerosName;
                                fols.Add(frString);

                            }
                            else if (fv)
                            {
                                fols.Add("");
                                fols.Add("Page");
                                string frString = "Folio " + noZerosName;
                                fols.Add(frString);
                            }
                            else
                            {
                                Console.WriteLine("ERROR: Folio outside of common DIPS string range. Investigate");
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                fols.Add(errString);
                                fols.Add("Page");
                                fols.Add(derivedFilename);
                            }

                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            fols.Add(fname);
                            fols.Add(errString);
                            fols.Add("Page");
                            fols.Add(derivedFilename);

                        }
                        folios.Add(fols);
                    }
                }
                // sort out numeric foliation here
                if (cNumericFolios.Any())
                {
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
                        if (match.Success)
                        {

                            nfols.Add(fname);
                            nfols.Add("");
                            nfols.Add("Page");
                            nfols.Add(derivedFilename);

                        }
                        else
                        {
                            Console.WriteLine("ERROR: Doesn't match numeric filenaming pattern");
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            nfols.Add(fname);
                            nfols.Add(errString);
                            nfols.Add("Page");
                            nfols.Add(derivedFilename);
                        }
                        numericFiles.Add(nfols);
                    }
                }

                if (cEndFlysheets.Any())
                {
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
                            efs.Add(fname);
                            var fser = Regex.Match(fname, @"(.)+((fse)[0-9]+[r])\.tif", RegexOptions.IgnoreCase).Success;
                            var fsev = Regex.Match(fname, @"(.)+((fse)[0-9]+[v])\.tif", RegexOptions.IgnoreCase).Success;
                            if (fser)
                            {
                                efs.Add(""); // error string
                                efs.Add("Flysheet");
                                string frString = "Back flysheet " + noZerosName;
                                efs.Add(frString);
                            }
                            else if (fsev)
                            {
                                efs.Add(""); // error string
                                efs.Add("Flysheet");
                                string frString = "Back flysheet " + noZerosName;
                                efs.Add(frString);
                            }
                            else
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                efs.Add(errString); // error string
                                efs.Add("Flysheet");
                                efs.Add(derivedFilename);
                            }
                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            efs.Add(fname);
                            efs.Add(errString);
                            efs.Add("Page");
                            efs.Add(derivedFilename);

                        }
                        endFlysheets.Add(efs);
                    }
                }
                if (cEndMatter.Any())
                {
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
                            ema.Add(fname);
                            var fbrigr = Regex.Match(fname, @"(.)+(fbrigr)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbrigv = Regex.Match(fname, @"(.)+(fbrigv)\.tif", RegexOptions.IgnoreCase).Success;
                            var fbspi = Regex.Match(fname, @"(.)+(fbspi)\.tif", RegexOptions.IgnoreCase).Success;

                            if (fbrigr)
                            {
                                ema.Add("");
                                ema.Add("Cover");
                                ema.Add("Back cover inside");
                            }
                            else if (fbrigv)
                            {
                                ema.Add("");
                                ema.Add("Cover");
                                ema.Add("Back cover");
                            }
                            else if (fbspi)
                            {
                                ema.Add("");
                                ema.Add("Cover");
                                ema.Add("Spine");
                            }
                            else // no match for any of these 'usual' cases
                            {
                                string errString = "Unexpected characters in filename. Flag for investigation";
                                ema.Add(errString);
                                ema.Add("Page");
                                ema.Add(derivedFilename);
                            }


                        }
                        else
                        {
                            string errString = "Unexpected characters in filename. Flag for investigation";
                            ema.Add(fname);
                            ema.Add(errString);
                            ema.Add("Page");
                            ema.Add(derivedFilename);
                        }
                        endMatter.Add(ema);
                    }
                }
                // check for anything else that passed through that failed the above checks


                frontMatter.Sort((a, b) => a[0].CompareTo(b[0]));
                folios.Sort((a, b) => a[0].CompareTo(b[0]));
                endFlysheets.Sort((a, b) => a[0].CompareTo(b[0]));
                endMatter.Sort((a, b) => a[0].CompareTo(b[0]));
                numericFiles.Sort((a, b) => a[0].CompareTo(b[0]));



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

                    Console.WriteLine("Mixture of DIPS-compliant and non-compliant filenames in shelfmark {0}", folderDerivedShelfmark);
                    // Just write this to console for now! Also would write to sharepoint for this shelfmark

                }

                foreach (List<String> fmList in frontMatter)
                {
                    allFilesSorted.Add(fmList);
                }
                foreach (List<String> folList in folios)
                {
                    allFilesSorted.Add(folList);
                }
                // Add in the numerically labelled files if they exist, we've sorted out the error flags here anyway...
                foreach (List<String> numfolList in numericFiles)
                {
                    allFilesSorted.Add(numfolList);
                }

                foreach (List<String> fsList in endFlysheets)
                {
                    allFilesSorted.Add(fsList);
                }
                foreach (List<String> emList in endMatter)
                {
                    allFilesSorted.Add(emList);
                }


            } // if at least some DIPS-compliant filenames exist
            else // is fully non-DIPS compliant and just has numerical filenames, so just sort this normally
            {
               List<String> sortedFilenames = fileNames.OrderBy(x => x).Select(x => x.ToString()).ToList();
                foreach (var sfn in sortedFilenames)
                {
                    List<String> nums = new List<String>();
                    nums.Add(sfn);
                    nums.Add(""); // errorString
                    nums.Add("Page");
                    string[] split = sfn.Split('.');
                    string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                    string fileExtension = split.Last(); // tif
                    string[] split2 = shelfmark_filename.Split('_');
                    string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                    string derivedFilename = split2.Last();
                    string noZerosName = derivedFilename.TrimStart('0');
                    noZerosName = noZerosName.Length > 0 ? noZerosName : "0";
                    nums.Add(noZerosName); // just get the number from the filename

                    allFilesSorted.Add(nums);
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
                allFilesSorted[i].Add(orderNumber);
            }

           // now allFilesSorted contains 
           //[0]: filename
           //[1]: flagStatus
           //[2]: objectType
           //[3]: label
           //[4]: order number
            return allFilesSorted;
        }

        private static bool writeFileLabelsToCSV(List<List<String>> ShelfmarkFilesLabels, String outFolder)
        {
            bool fError = false;

            try // to write the csv...
            {
                const char sep = ',';
                List<String> strHeaders = new List<string>{"File","Order","Type","Label"};
                System.Text.UnicodeEncoding uce = new System.Text.UnicodeEncoding();
                string fNameString = "ImageOrder";
                string outPath = outFolder + @"\"+fNameString+".csv";

                if (File.Exists(outPath))
                {
                    var time = DateTime.Now;
                    string formattedTime = time.ToString("yyyyMMdd_hh-mm-ss");
                    string altOutPath = outFolder + @"\" + fNameString + "_" + formattedTime + ".csv";
                    outPath = altOutPath;

                }
                    

                using (var sr = new StreamWriter(outPath, false, uce))
                {
                    using (var csvFile = new CsvHelper.CsvWriter(sr, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        foreach (var header in strHeaders)
                        {
                            csvFile.WriteField(header);
                        }
                        csvFile.NextRecord(); // skips to next line...
                        foreach (var record in ShelfmarkFilesLabels)
                        { 
                            csvFile.WriteField(record[0]); // filename
                            csvFile.WriteField(record[4]); // order number
                            csvFile.WriteField(record[2]); // object type
                            csvFile.WriteField(record[3]); // label
                            csvFile.WriteField(record[1]); // error flag status
                        
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


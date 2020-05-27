using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    class InputOrderSpreadsheetTools
    {
        public static List<List<String>> getAllShelfmarkTIFs(List<List<String>> sharepointOut)
        {
            bool fError = false;
            List<List<String>> sourceFolderXMLs = new List<List<String>>(); // maybe don't need?
            List<List<String>> allShelfmarkTIFs = new List<List<String>>();


            for (int i = 1; i < sharepointOut.Count; i++) // need this to skip the first row (titles)
            {
                List<String> item = sharepointOut[i];
                List<String> shelfmarkTIFs = new List<String>();

                bool validPath = false;
                var shelfmark = item[1];

                if (item[5] != "false")
                {
                    var testFullName = item[6];
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

                    if (sourceFolder.EndsWith("TIF"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (sourceFolder.EndsWith("tif"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (sourceFolder.EndsWith("tiff"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (sourceFolder.EndsWith("TIFF"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (sourceFolder.EndsWith("tiffs"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (sourceFolder.EndsWith("TIFFS"))
                    {
                        tifFolder = sourceFolder;
                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "TIFF"))
                    {
                        tifFolder = sourceFolder + @"\" + "TIFF";

                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "tiff"))
                    {
                        tifFolder = sourceFolder + @"\" + "tiff";

                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "TIFFS"))
                    {
                        tifFolder = sourceFolder + @"\" + "TIFFS";

                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "tiffs"))
                    {
                        tifFolder = sourceFolder + @"\" + "tiffs";

                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "tif"))
                    {
                        tifFolder = sourceFolder + @"\" + "tif";

                    }
                    else if (Directory.Exists(sourceFolder + @"\" + "TIF"))
                    {
                        tifFolder = sourceFolder + @"\" + "TIF";

                    }
                    else
                    {
                        Console.WriteLine("No folder found for shelfmark {0}", shelfmark);
                    }

                    if (Directory.Exists(tifFolder))
                    {
                        validPath = true;
                    }


                    // now got the tiff folder, need to check the list of files that appears

                    // first check it exists:
                    if (validPath)
                    {

                        DirectoryInfo d = new DirectoryInfo(tifFolder);
                        FileInfo[] Files = d.GetFiles("*.TIF*");

                        // Can then add this to a list of strings
                        string str = "";
                        var numberOfItems = Files.Length; // only do this once per shelfmark
                                                          // do you need this?

                        shelfmarkTIFs.Add(shelfmark);

                        List<List<String>> shelfmarkLabels = mapFileNameToLabels(Files);



                    }// if validPath == true


                    else // so not a valid path!
                    {
                        shelfmarkTIFs.Add(shelfmark);
                        shelfmarkTIFs.Add(null);
                        continue; // use continue for now, but will need to write out invalid path to a variable at some point
                    }
                }
                else
                {
                    shelfmarkTIFs.Add(shelfmark);
                    shelfmarkTIFs.Add(null);
                    // Got yourself a shelfmark that needs checking, so obviously things will fail here...
                    continue;
                }
                allShelfmarkTIFs.Add(shelfmarkTIFs);
            } // end of the for loop over each shelfmark


            return allShelfmarkTIFs;
        }
        private static List<List<string>> mapFileNameToLabels(FileInfo[] Files)
        {

            // Order labels will take a couple of sweeps - one to get front and back matter and then another to do a fine sort of the front and back matter
            List<String> shelfmarkLabels = new List<String>();

            Dictionary<string, string> order_map = new Dictionary<string, string>();
            List<string> fileNames = Files.Select(x => x.Name).ToList();


            // List<String> frontMatter = new List<String>();
            // List<String> endMatter = new List<String>();
            // List<String> folios = new List<String>();

            List<String> frontNames = new List<String>();
            List<String> endNames = new List<String>();

            // Make front covers appear first in the order
            order_map.Add("fblefr", "1");
            order_map.Add("fblefv", "2");
            order_map.Add("fbspi", Files.Length.ToString()); // spine is always the last item?
            order_map.Add("fbrigv", ((Files.Length) - 1).ToString()); // back cover
            order_map.Add("fbrigr", ((Files.Length) - 2).ToString()); // back cover inside

            // front matter
            frontNames.Add("fblefr");
            frontNames.Add("fblefv");
            frontNames.Add("fs"); // adding front flysheets

            // end matter
            endNames.Add("fbspi"); // spine is always the last item?
            endNames.Add("fbrigv"); // back cover
            endNames.Add("fbrigr"); // back cover inside
            endNames.Add("fse"); // end flysheets

            // Define regular expressions to search for
            // 'initial' versions perform a looser search
            var frontMatterRegex = new Regex(@"(.)+(((fble)((fv)|(fr)))|((fs)[0-9]+[rv]))\.tif", RegexOptions.IgnoreCase);
            var initialFrontMatterRegex = new Regex(@"(.)+(((fble)((fv)|(fr)))|((fs)[0-9]+(.)+))\.tif", RegexOptions.IgnoreCase);
            
            var folioRegex = new Regex(@"(.)+(f)([0-9])+([rv])\.tif", RegexOptions.IgnoreCase);
            var initialFolioRegex = new Regex(@"(.)+(f)([0-9])+(.)+\.tif", RegexOptions.IgnoreCase);

            var endFlysheetsRegex = new Regex(@"(.)+((fse)[0-9] +[rv])\.tif", RegexOptions.IgnoreCase);
            var initialEndFlysheetsRegex = new Regex(@"(.)+((fse)[0-9]+(.)+)\.tif", RegexOptions.IgnoreCase);

            var endMatterRegex = new Regex(@"(((fb)((rigv)|(rigr)|(spi))))\.tif", RegexOptions.IgnoreCase);
            var initialEndMatterRegex = new Regex(@"(((fb)((rigv)|(rigr)|(spi))))\.tif", RegexOptions.IgnoreCase);
            // Sort into front matter, end flysheets, end matter and folios

            // Candidates for each section:
            List<string> cFrontMatter = fileNames.Where(f => initialFrontMatterRegex.IsMatch(f)).ToList();
            List<string> cEndFlysheets = fileNames.Where(f => initialEndFlysheetsRegex.IsMatch(f)).ToList();
            List<string> cEndMatter = fileNames.Where(f => initialEndMatterRegex.IsMatch(f)).ToList();
            List<string> cFolios = fileNames.Where(f => initialFolioRegex.IsMatch(f)).ToList();

            List<List<String>> allFilesSorted = new List<List<String>>();
            List<List<string>> frontMatter = new List<List<String>>();
            List<List<string>> endFlysheets = new List<List<String>>();
            List<List<string>> endMatter = new List<List<String>>();
            List<List<string>> folios = new List<List<String>>();

            if (cFrontMatter.Any() | cFolios.Any() | cEndMatter.Any() | cEndFlysheets.Any())
            {
                // you can be pretty sure its DIPs compliant if you see any titles or any numbered folios
                foreach (string fname in cFrontMatter)
                {
                    List<String> fmat = new List<String>();
                    var match = Regex.Match(fname, @"(.)+(((fble)((fv)|(fr)))|((fs)[0-9]+[rv]))\.tif",RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        fmat.Add(fname);
                        fmat.Add("");

                    }
                    else
                    {
                        string errString = "Unexpected characters in filename. Flag for investigation";
                        fmat.Add(fname);
                        fmat.Add(errString);
                    }
                    frontMatter.Add(fmat);
                }
                foreach (string fname in cFolios)
                {
                    List<String> fols = new List<String>();
                    var match = Regex.Match(fname, @"(.)+(f)([0-9])+([rv])\.tif", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        fols.Add(fname);
                        fols.Add("");

                    }
                    else
                    {
                        string errString = "Unexpected characters in filename. Flag for investigation";
                        fols.Add(fname);
                        fols.Add(errString);
                    }
                    folios.Add(fols);
                }
                foreach (string fname in cEndFlysheets)
                {
                    List<String> efs = new List<String>();
                    var match = Regex.Match(fname, @"(.)+((fse)[0-9] +[rv])\.tif", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        efs.Add(fname);
                        efs.Add("");

                    }
                    else
                    {
                        string errString = "Unexpected characters in filename. Flag for investigation";
                        efs.Add(fname);
                        efs.Add(errString);
                    }
                    endFlysheets.Add(efs);
                }
                foreach (string fname in cEndMatter)
                {
                    List<String> ema = new List<String>();
                    var match = Regex.Match(fname, @"(((fb)((rigv)|(rigr)|(spi))))\.tif", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        ema.Add(fname);
                        ema.Add("");

                    }
                    else
                    {
                        string errString = "Unexpected characters in filename. Flag for investigation";
                        ema.Add(fname);
                        ema.Add(errString);
                    }
                    endMatter.Add(ema);
                }
                // check for anything else that passed through that failed the above checks

                frontMatter.Sort((a, b) => a[0].CompareTo(b[0]));
                folios.Sort((a, b) => a[0].CompareTo(b[0]));
                endFlysheets.Sort((a, b) => a[0].CompareTo(b[0]));
                endMatter.Sort((a, b) => a[0].CompareTo(b[0]));



                foreach (List<String> fmList in frontMatter)
                {
                    allFilesSorted.Add(fmList);
                }
                foreach (List<String> fmList in folios)
                {
                    allFilesSorted.Add(fmList);
                }
                foreach (List<String> fmList in endFlysheets)
                {
                    allFilesSorted.Add(fmList);
                }
                foreach (List<String> fmList in endMatter)
                {
                    allFilesSorted.Add(fmList);
                }


            } // if DIPs compliant
            else
            {
               List<String> sortedFilenames = fileNames.OrderBy(x => x).Select(x => x.ToString()).ToList();
                foreach (var sfn in sortedFilenames)
                {
                    List<String> nums = new List<String>();
                    nums.Add(sfn);
                    nums.Add("");
                    allFilesSorted.Add(nums);
                }
            }

            /*


            List<string> frontMatter = fileNames.Where(f => frontMatterRegex.IsMatch(f)).ToList();
            List<string> endFlysheets = fileNames.Where(f => endFlysheetsRegex.IsMatch(f)).ToList();
            List<string> endMatter = fileNames.Where(f => endMatterRegex.IsMatch(f)).ToList();
            List<string> folios = fileNames.Where(f => folioRegex.IsMatch(f)).ToList();

            List<String> allFilesSorted = new List<String>();
            if (frontMatter.Any() | folios.Any()) // is DIPS format
            {
                List<String> sortedFrontMatter = frontMatter.OrderBy(x => x).ToList();
                List<String> sortedEndMatter = endMatter.OrderBy(x => x).ToList();
                List<String> sortedEndFlysheets = endFlysheets.OrderBy(x => x).ToList();
                List<String> sortedFolios = folios.OrderBy(x => x).ToList();
                allFilesSorted = sortedFrontMatter.Concat(sortedFolios).Concat(sortedEndFlysheets).Concat(sortedEndMatter).ToList();
            }
            else
            {
                allFilesSorted = fileNames.OrderBy(x => x).ToList();
            }
            int counter = 1;
            
            */

            /*
            foreach (var file in Files)
            {
                string order = "";
                string type = "";
                string label = "";
                bool flag = false;

                string s = file.Name;
                shelfmarkLabels.Add(s);
                Console.WriteLine(s);
                string[] split = s.Split('.');
                string shelfmark_filename = string.Join(".", split.Take(split.Length - 1)); // shelfmark_filename
                string fileExtension = split.Last(); // tif
                string[] split2 = shelfmark_filename.Split('_');
                string derivedShelfmark = string.Join(".", split2.Take(split2.Length - 1)); // shelfmark
                string derivedFilename = split2.Last();

                // Now need to build a map of file names to file labels

                bool isEndMatter = endNames.Any(x => derivedFilename.Contains(x));
                bool isFrontMatter = frontNames.Any(x => derivedFilename.Contains(x));

                if (isEndMatter)
                {

                    endMatter.Add(file.Name);


                }
                else if (isFrontMatter)
                {
                    frontMatter.Add(file.Name)
                }
                else // should just be folios left, but check
                {
                    var folioRegex = new Regex(@"(f)([0-9])+([rv])\.tif$",RegexOptions.IgnoreCase);
                    List<string> resultList = derivedFilename.Where(folioRegex.IsMatch).ToList();

                }

                // Use a separate function
                foreach (KeyValuePair<string, string> entry in order_map) // check if any matches from dict
                {

                    if (derivedFilename.Contains(entry.Key))
                    {
                        order = entry.Value;
                    }

                } // end of dict loop

                shelfmarkLabels.Add(order);
                shelfmarkLabels.Add(type);
                shelfmarkLabels.Add(label);

            }
        */

            return allFilesSorted;
        }


        public static bool RetrieveImgOrderLabels(List<List<String>> allShelfmarkFiles)
        {
            bool fError = false;

            foreach (List<String> shelfmarkFiles in allShelfmarkFiles)
            {
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                filePath += @"\HMDSharepoint_ImgOrderTest\";
                string SM_folderFormat = shelfmarkFiles[0].ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*", @"~");
                filePath += SM_folderFormat;

                if (!Directory.Exists(filePath))
                {


                    Directory.CreateDirectory(filePath);
                }

                foreach (var thing in shelfmarkFiles)
                {
                    // Create a directory on the user's desktop if it doesn't already exist

                    Console.WriteLine("{0}", thing);
                }
            }


            return !fError;
        }
    }
}


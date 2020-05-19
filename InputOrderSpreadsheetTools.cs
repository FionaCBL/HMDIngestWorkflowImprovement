using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace HMDSharepointChecker
{
    class InputOrderSpreadsheetTools
    {
        public static List<List<String>> getAllShelfmarkTIFs(List<List<String>> sharepointOut)
        {
            bool fError = false;
            List<List<String>> sourceFolderXMLs = new List<List<String>>(); // maybe don't need?
            List<List<String>> allShelfmarkTIFs = new List<List<String>>(); // maybe don't need?


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

                        foreach (FileInfo file in Files)
                        {
                            string imgOrder_Order = "";
                            string imgOrder_Type = "";
                            string imgOrder_Label = "";

                            String name = file.Name;

                            String[] sep = { ".tif" };
                            Int32 count = 1;

                            shelfmarkTIFs.Add(name);

                            // using the method 
                            String[] strlist = name.Split(sep, count,
                                   StringSplitOptions.RemoveEmptyEntries);

                            foreach (String s in strlist)
                            {
                                Console.WriteLine(s);
                            }
                        }
                    } // if validPath == true

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

        public static bool RetrieveImgOrderLabels(List<List<String>> allShelfmarkFiles)
        {
            bool fError = false;

            foreach (List<String> shelfmarkFiles in allShelfmarkFiles)
            {
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                filePath += @"\HMDSharepoint_ImgOrderTest\";
                string SM_folderFormat = shelfmarkFiles[0].ToLower().Replace(@" ", @"_").Replace(@"/", @"!").Replace(@".", @"_").Replace(@"*",@"~");
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


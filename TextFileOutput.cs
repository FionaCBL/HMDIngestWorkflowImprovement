using System;
using System.IO;
using System.Collections.Generic;

namespace HMDSharepointChecker
{
    class TextOutputFunctions
    {
        public static bool OutputListOfLists(List<HMDObject> HMDOut, string suffix) 
        {

            bool fError = false;
            // These examples assume a "C:\Users\Public\TestFolder" folder on your machine.
            // You can modify the path if necessary.

            try
            {
                var dayNow = DateTime.Now.ToString("dd_MM_yy"); // includes leading zeros
                var timeNow = DateTime.Now.ToString("HH_mm_ss");
                string filename = null;


                // Create a directory on the user's desktop if it doesn't already exist
                string filePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                filePath += @"\HMDSharepointLogs\";
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }


                if (!string.IsNullOrEmpty(suffix))
                {
                    filename = filePath+@"\HMDSharepointOutput_" + suffix + "_" + dayNow + "_" + timeNow + ".txt"; // this is hardcoded for now, clean this up going fwd

                }
                else
                {
                    filename = filePath+@"\HMDSharepointOutput_" + dayNow + "_" + timeNow + ".txt"; // this is hardcoded for now, clean this up going fwd

                }
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(filename))
                {
                    foreach (var line in HMDOut)
                    {
                        file.WriteLine(string.Join(",",line));

                    }
                }

                if (!File.Exists(filename))
                {
                    fError = true;
                }
            }
            catch
            {
                fError = true;
            }

            
            return !fError;
        }
    }
}
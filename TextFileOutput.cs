using System;
using System.IO;
using System.Collections.Generic;

namespace HMDSharepointChecker
{
    class TextOutputFunctions
    {
        public static bool OutputListOfLists(List<List<String>> HMDOut) 
        {

            bool fError = false;
            // These examples assume a "C:\Users\Public\TestFolder" folder on your machine.
            // You can modify the path if necessary.

            try
            {
                var dayNow = DateTime.Now.ToString("dd_MM_yy"); // includes leading zeros
                var timeNow = DateTime.Now.ToString("HH_mm_ss");

                var filename = @"C:\Users\hjmoss\Desktop\HMDSharepointOutput_" + dayNow + "_"+ timeNow+".txt"; // this is hardcoded for now, clean this up going fwd
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(filename))
                {
                    foreach (List<String> line in HMDOut)
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
//Output (to WriteLines.txt):
//   First line
//   Second line
//   Third line

//Output (to WriteText.txt):
//   A class is the most powerful data type in C#. Like a structure, a class defines the data and behavior of the data type.

//Output to WriteLines2.txt after Example #3:
//   First line
//   Third line

//Output to WriteLines2.txt after Example #4:
//   First line
//   Third line
//   Fourth line
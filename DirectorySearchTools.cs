using System;
using System.IO;
using System.Collections.Generic;

namespace HMDSharepointChecker
{
    class DirectorySearchTools
    {

        public static String[] GetFilesFrom(String searchFolder, String[] filters, bool isRecursive)
        {
            // Use like:
            //String searchFolder = @"C:\MyFolderWithImages";
            //var filters = new String[] { "jpg", "jpeg", "png", "gif", "tiff", "bmp", "svg" };
            //var files = GetFilesFrom(searchFolder, filters, false);


            List<String> filesFound = new List<String>();
            var searchOption = isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            foreach (var filter in filters)
            {
                filesFound.AddRange(Directory.GetFiles(searchFolder, String.Format("*.{0}", filter), searchOption));
            }
            return filesFound.ToArray();
        }



    }
}



using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace Affirma.OpenSource.PresentationTool
{
    public static class Utils
    {
        /// <summary>
        /// Creates a temporary file and renames it to .pptx file format
        /// </summary>
        /// <returns></returns>
        public static string GetTempFile()
        {
            var fileName = Path.GetTempFileName();
            var presentationFile = $"{fileName}.pptx";
            File.Move(fileName, presentationFile);
            return presentationFile;
        }

        /// <summary>
        /// Utility function to check validity of presentation files
        /// </summary>
        /// <param name="documentURLs"></param>
        public static void CheckPresentationFiles(IList<string> documentURLs)
        {
            foreach(string url in documentURLs)
            {
                if(Path.GetExtension(url) != ".pptx")
                {
                    throw new ArgumentException($"{url} does not have a valid presentation format. Supported formats include: {".pptx"}");
                }
                if(!File.Exists(url))
                {
                    throw new ArgumentException($"Cannot find file '{url}'");
                }
            }
        }
    }
}

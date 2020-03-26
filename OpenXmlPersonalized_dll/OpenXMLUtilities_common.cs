#region Riferimenti
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
#endregion Riferimenti

namespace OpenXmlPersonalized_dll
{
    public class OpenXMLUtilities_common
    {
        //Percorso base in bin debug
        public static string createPath(string fileExtension, string OutputFileDirectory = @".\", string fileName = "Output")
        {
            var datetime = DateTime.Now.ToString().Replace("/", "_").Replace(":", "_");

            string fileFullname = Path.Combine(OutputFileDirectory, $"{fileName}.{fileExtension}");

            if (File.Exists(fileFullname))
            {
                fileFullname = Path.Combine(OutputFileDirectory, $"{fileName}_{datetime}.{fileExtension}");
            }
            return fileFullname;
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CefSharp.WinForms;

namespace Dax.Scrapping.Appraisal.Core
{
    public static class ChromiumBrowserExtension
    {
        /// <summary>
        /// Extension method to read a js file into Scripts folder and return it as string.
        /// </summary>
        /// <param name="browser">Instanced browser object</param>
        /// <param name="filename">Js filename</param>
        /// <returns></returns>
        public static string GetScriptText(this ChromiumWebBrowser browser, string filename)
        {
            string scriptPath = Path.GetFullPath(Directory.GetCurrentDirectory() + @"\Scripts\" + filename);

            try
            {   // Open the text file using a stream reader.
                using (StreamReader sr = new StreamReader(scriptPath))
                {
                    // Read the stream to a string.
                    String script = sr.ReadToEnd();
                    return script;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }
            return string.Empty;
        }
    }
}

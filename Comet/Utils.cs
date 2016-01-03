using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TsudaKageyu;

namespace Comet
{
    public static class Utils
    {

        public static Image GetIcon(String iconPath)
        {
            if (String.IsNullOrEmpty(iconPath) || !File.Exists(iconPath))
            {
                return null;
            }

            var ext = Path.GetExtension(iconPath);
            if (ext == ".exe") // Also works with .dll
            {
                /* IconExtractor does a better job (produces better resolution) extracting
                 * icons from .exe/.dll files then Icon.ExtractAssociatedIcon()
                 * Example was taken from here: 
                 * http://www.codeproject.com/Articles/26824/Extract-icons-from-EXE-or-DLL-files
                 */
                var iconExtractor = new IconExtractor(iconPath);
                try
                {
                    var icon = iconExtractor.GetIcon(0);
                    var splitIcons = IconUtil.Split(icon);
                    return IconUtil.ToBitmap(splitIcons[0]);
                }
                catch (Exception)
                {
                    return null;
                }

            }

            if (ext == ".ico")
            {
                var icon = Icon.ExtractAssociatedIcon(iconPath);
                return icon == null ? null : icon.ToBitmap();
            }

            if (ext == ".png" || ext == ".bmp" || ext == ".gif" ||
                ext == ".jpg" || ext == ".jpeg" || ext == ".tiff")
            {
                return Image.FromFile(iconPath);
            }

            return null;
        }

        public static void UpdateEnvironmentVariables()
        {
            /* A regular user environment variable overrides completely a system one with
             * the same name if both exist, but only for the specific user it is specified 
             * for. However, the user path variables is treated differently. It is appended 
             * to the system path variable when evaluating, rather than completely 
             * replacing it. Source:
             * http://stackoverflow.com/questions/5126512/how-environment-variables-are-evaluated
             */

            var sysEnvVars = Environment.GetEnvironmentVariables(EnvironmentVariableTarget.Machine);
            foreach (DictionaryEntry envVar in sysEnvVars)
            {
                Environment.SetEnvironmentVariable((String)envVar.Key, (String)envVar.Value);
            }

            var usrEnvVars = Environment.GetEnvironmentVariables(EnvironmentVariableTarget.User);
            foreach (DictionaryEntry envVar in usrEnvVars)
            {
                // The PATH variable is treated differently
                if ((String)envVar.Key == "PATH")
                {
                    String sysPath = Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine);
                    String combinedPath = sysPath + ";" + (String)envVar.Value; // Combine system and user paths
                    Environment.SetEnvironmentVariable("PATH", combinedPath);
                    continue;
                }
                Environment.SetEnvironmentVariable((String)envVar.Key, (String)envVar.Value);
            }
        }

        public static string GetStandardExecutablePath(string fileName)
        {
            // Expand all windows environment %variables%
            fileName = Environment.ExpandEnvironmentVariables(fileName);

            // If full path was specified from the beginning
            if (File.Exists(fileName))
            {
                return Path.GetFullPath(fileName);
            }

            // Check to see if path has extension
            var extensionList = new List<String>();
            var fileExt = Path.GetExtension(fileName);
            if (fileExt != String.Empty)
            {
                extensionList.Add(fileExt);
                fileName = Path.GetFileNameWithoutExtension(fileName);
                // Catch cases with filenames consisting of extension only
                if (fileName == "")
                {
                    return null;
                }
            }
            else
            {
                extensionList = (Environment.GetEnvironmentVariable("PATHEXT") ?? "").Split(';').ToList();
                if (extensionList.Count == 0) // Extension list is empty
                {
                    return null;
                }
            }

            // Get PATH variable from the current process
            var pathList = (Environment.GetEnvironmentVariable("PATH") ?? "").ToLower().Split(';');
            if (pathList.Length == 0) // Path list is empty
            {
                return null;
            }

            foreach (var path in pathList)
            {
                foreach (var ext in extensionList)
                {
                    var fullPath = Path.Combine(path, fileName) + ext;
                    if (File.Exists(fullPath))
                        return fullPath;
                }
            }

            return null;
        }
    }
}

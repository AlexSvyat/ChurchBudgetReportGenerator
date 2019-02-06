using System.IO;

namespace ChurchBudgetReportGenerator
{
    /// <summary>
    /// Various File Utils used in this project
    /// </summary>
    public class Utils
    {
        static DirectoryInfo _outputDir = null;
        public static DirectoryInfo OutputDir
        {
            get
            {
                return _outputDir;
            }
            set
            {
                _outputDir = value;
                if (!_outputDir.Exists)
                {
                    _outputDir.Create();
                }
            }
        }

        public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
        {
            if (!File.Exists(file))
            {
                file = OutputDir.FullName + Path.DirectorySeparatorChar + file;
            }
            var fi = new FileInfo(file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
    }
}
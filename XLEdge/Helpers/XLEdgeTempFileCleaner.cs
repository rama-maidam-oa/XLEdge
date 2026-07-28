using System;
using System.IO;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Clears out leftover temporary report CSV files (see ReportGenerator.WriteTempCsv) from
    /// XLEdgeAppPaths.TempFolder, both once at add-in startup and after every completed report run,
    /// so temp files don't silently accumulate. Files still open/locked are skipped rather than
    /// risking a delete failure.
    /// </summary>
    public static class XLEdgeTempFileCleaner
    {
        public static void DeleteAllTempFiles()
        {
            const string MethodName = nameof(DeleteAllTempFiles);

            try
            {
                string tempFolder = XLEdgeAppPaths.TempFolder;
                var directoryInfo = new DirectoryInfo(tempFolder);

                if (!directoryInfo.Exists)
                {
                    try
                    {
                        directoryInfo.Create();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"{MethodName}|Failed to create temp folder {tempFolder}");
                    }

                    return;
                }

                foreach (FileInfo file in directoryInfo.GetFiles())
                {
                    try
                    {
                        if (IsFileOpen(file.FullName))
                        {
                            continue;
                        }

                        file.Delete();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"{MethodName}|Exception in deleting file {file.FullName}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Exception in deleting files.");
            }
        }

        /// <summary>
        /// Attempts an exclusive read/write open; if that fails, the file is presumed still in use
        /// and is skipped rather than risking a delete failure mid-cleanup.
        /// </summary>
        private static bool IsFileOpen(string fileName)
        {
            const string MethodName = nameof(IsFileOpen);

            if (!File.Exists(fileName))
            {
                return false;
            }

            try
            {
                using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // Successfully opened exclusively - not in use by anyone else.
                }

                return false;
            }
            catch (IOException ex)
            {
                LogUtility.LogWarn($"{MethodName}|File {fileName} is in use. Skipping deletion. {ex.Message}");
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Unexpected error checking if file {fileName} is open");
                return true;
            }
        }
    }
}

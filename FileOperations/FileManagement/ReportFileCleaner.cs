using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.FileOperations.FileManagement
{
    /// <summary>
    /// Responsible for cleaning up files and directories based on specified rules.
    /// </summary
    public static class ReportFileCleaner
    {
        /// <summary>
        /// Remove reports.
        /// </summary>
        /// <param name="destinationFolder">
        /// Path to folder with reports.
        /// </param>
        /// <param name="reportFileName">
        /// Default report file name.
        /// </param>
        /// <param name="singleFile">
        /// Indicates whether reports should be in single file.
        /// </param>
        /// <param name="reportCleanup">
        /// Indicates whether all reports are to be deleted.
        /// </param>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        public static void Removereports(
            string destinationFolder,
            string reportFileName,
            bool singleFile,
            bool reportCleanup,
            ILogger logger)
        {
            using (logger.BeginScope("Remove reports"))
            {
                if (Directory.Exists(destinationFolder))
                {
                    if (reportCleanup)
                    {
                        logger.LogInformation($"Deleting recursive {destinationFolder}");
                        try
                        {
                            Directory.Delete(destinationFolder, true);
                        }
                        catch (IOException ioExc)
                        {
                            logger.LogCritical(ioExc.Message);
                        }
                    }
                    else if (singleFile)
                    {
                        string[] files = Directory.GetFiles(
                            destinationFolder,
                            string.Format("{0}*", reportFileName));
                        foreach (string file in files)
                        {
                            logger.LogInformation($"Deleting {file}");
                            File.Delete(file);
                        }
                    }
                }
                else
                {
                    logger.LogWarning("There is no {0}, nothing to remove!", destinationFolder);
                }

                logger.LogInformation("Deletion completed");
            }
        }
    }
}

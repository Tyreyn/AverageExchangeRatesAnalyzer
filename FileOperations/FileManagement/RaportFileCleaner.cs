using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.FileOperations.FileManagement
{
    public static class RaportFileCleaner
    {
        /// <summary>
        /// Remove raports.
        /// </summary>
        /// <param name="DestinationFolder">
        /// Path to folder with raports.
        /// </param>
        /// <param name="RaportFileName">
        /// Default raport file name.
        /// </param>
        /// <param name="SingleFile">
        /// Indicates whether reports should be in single file.
        /// </param>
        public static void RemoveRaports(
            string DestinationFolder,
            string RaportFileName,
            bool SingleFile,
            bool RaportCleanup,
            ILogger logger)
        {
            using (logger.BeginScope("Remove Raports"))
            {
                if (Directory.Exists(DestinationFolder))
                {
                    if (RaportCleanup)
                    {
                        logger.LogInformation($"Deleting recursive {DestinationFolder}");
                        Directory.Delete(DestinationFolder, true);
                    }
                    else if (SingleFile)
                    {
                        string[] files = Directory.GetFiles(DestinationFolder,
                                                            searchPattern: String.Format("{0}*", RaportFileName));
                        foreach (string file in files)
                        {
                            logger.LogInformation($"Deleting {file}");
                            File.Delete(file);
                        }
                    }
                }
                else
                {
                    logger.LogWarning("There is no {0}, nothing to remove!", DestinationFolder);
                }
                logger.LogInformation("Deletion completed");
            }

        }
    }
}

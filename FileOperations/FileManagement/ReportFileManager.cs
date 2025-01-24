using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AverageExchangeRatesAnalyzer.FileOperations.FileManagement
{
    /// <summary>
    /// Responsible for moving files.
    /// </summary
    public class ReportFileManager
    {
        /// <summary>
        /// Logger class.
        /// </summary>
        private readonly ILogger? logger;

        /// <summary>
        /// Configuration root.
        /// </summary>
        private readonly IConfigurationRoot configuration;

        /// <summary>
        /// Day data was downloaded.
        /// </summary>
        private readonly DateTime currentDate;

        /// <summary>
        /// Default name of exported file.
        /// </summary>
        private string reportFileName = "Currency_rate_report";

        /// <summary>
        /// Default name of reports folder.
        /// </summary>
        private string reportFolderName = "reports";

        /// <summary>
        /// Temporary name of report file.
        /// </summary>
        private string tmpReportFileName = "tmp.xlsx";

        /// <summary>
        /// Destination folder where report will be stored.
        /// </summary>
        private string destinationFolder = Directory.GetCurrentDirectory();

        /// <summary>
        /// Initializes a new instance of the <see cref="ReportFileManager"/> class.
        /// </summary>
        /// <param name="logger">
        /// Logger class.
        /// </param>
        /// <param name="configuration">
        /// Configuration root.
        /// </param>
        /// <param name="currentDate">
        /// Day data was downloaded.
        /// </param>
        public ReportFileManager(ILogger logger, IConfigurationRoot configuration, DateTime currentDate)
        {
            this.logger = logger;
            this.configuration = configuration;
            this.currentDate = currentDate;
            this.logger.LogInformation("Initializing ReportFileManager");
            this.LoadSettings();
        }

        /// <summary>
        /// Gets or sets a value indicates whether reports should be in single file.
        /// </summary>
        private bool SingleFile { get; set; } = false;

        /// <summary>
        /// Gets or sets a value indicates whether all reports are to be deleted.
        /// </summary>
        private bool ReportCleanup { get; set; } = false;

        /// <summary>
        /// Perform folder clean, files moving.
        /// </summary>
        public void PerformManagementOfNewReport()
        {
            using (this.logger.BeginScope("Perform management of a new report"))
            {
                string pathToreports = Path.Combine(this.destinationFolder, this.reportFolderName);
                ReportFileCleaner.Removereports(
                    pathToreports,
                    this.reportFolderName,
                    this.SingleFile,
                    this.ReportCleanup,
                    this.logger);
                Directory.CreateDirectory(pathToreports);
                if (File.Exists(this.tmpReportFileName))
                {
                    if (this.SingleFile)
                    {
                        try
                        {
                            File.Move(
                                Path.Combine(this.tmpReportFileName),
                                Path.Combine(pathToreports, string.Format("{0}_{1}.xlsx", this.reportFileName, this.currentDate.ToString("ddMMyyyy"))));
                        }
                        catch (IOException ioExc)
                        {
                            this.logger.LogCritical(ioExc.Message);
                        }
                    }
                    else
                    {
                        string separateFolder = Path.Combine(pathToreports, this.currentDate.ToString("yyyy_MM"));
                        if (!Directory.Exists(separateFolder))
                        {
                            Directory.CreateDirectory(separateFolder);
                        }

                        try
                        {
                            File.Move(
                                Path.Combine(this.tmpReportFileName),
                                Path.Combine(separateFolder, string.Format("{0}_{1}.xlsx", this.reportFileName, this.currentDate.ToString("ddMMyyyy"))));
                            this.logger.LogInformation(
                                "report saved in {0}",
                                Path.Combine(separateFolder, string.Format("{0}_{1}.xlsx", this.reportFileName, this.currentDate.ToString("ddMMyyyy"))));
                        }
                        catch (IOException ioExc)
                        {
                            this.logger.LogCritical(ioExc.Message);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Load required configuration.
        /// </summary>
        private void LoadSettings()
        {
            string tmpreportFolderName = this.configuration.GetValue<string>("reportFolderName");
            string tmpreportFileName = this.configuration.GetValue<string>("reportFileName");
            string tmpDestinationFolder = this.configuration.GetValue<string>("DestinationFolder");
            if (tmpDestinationFolder != null)
            {
                this.destinationFolder = tmpDestinationFolder;
            }

            this.SingleFile = this.configuration.GetValue<bool>("SingleFile");

            this.ReportCleanup = this.configuration.GetValue<bool>("reportCleanup");

            if (tmpreportFileName != null)
            {
                this.reportFileName = tmpreportFileName;
            }

            if (tmpreportFolderName != null)
            {
                this.reportFolderName = tmpreportFolderName;
            }

            this.logger?.LogInformation(
                "File name will be {FileName}_<end_date>",
                this.reportFileName);

            if (this.SingleFile)
            {
                this.logger?.LogInformation("Single File Mode");
                this.logger?.LogInformation(
                    "Files will be saved in {Destination}\\reports",
                    Path.Combine(this.destinationFolder, this.reportFolderName));
            }
            else
            {
                this.logger?.LogInformation(
                    "Files will be saved in {Destination}\\reports\\<month_name>",
                    Path.Combine(this.destinationFolder, this.reportFolderName));
            }

            if (this.ReportCleanup)
            {
                this.logger?.LogInformation("All reports will be deleted");
            }
        }
    }
}

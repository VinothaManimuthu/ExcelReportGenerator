using OfficeOpenXml;
using log4net;
using log4net.Config;
using System.Globalization;
namespace ExcelReportGenerator
{
    public class ReadDataFromExcel
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        /// <summary>
        /// Reads data from the specified Excel file and filters it based on the given state.
        /// </summary>
        /// <param name="inputFilePath">Path to the input Excel file.</param>
        /// <param name="stateFilter">State code to filter the data.</param>
        /// <returns>A list of DataRecord objects containing the filtered data.</returns>
        public static List<DataRecord> ReadData(string inputFilePath, string stateFilter)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Configure log4net logging
            XmlConfigurator.Configure(new FileInfo("log4net.config"));
            log.Info("Reading excel data for the state " + stateFilter);
            List<DataRecord> dataRecords = new List<DataRecord>();
            //  Dictionary to store Refer and Admit count based on date
            Dictionary<DateOnly, int> referCountByDate = new Dictionary<DateOnly, int>();
            Dictionary<DateOnly, int> admitCountByDate = new Dictionary<DateOnly, int>();

            // Check if the file exists
            if (!File.Exists(inputFilePath))
            {
                return dataRecords; // Return empty list if file doesn't exist
            }

            // Open the input Excel file
            using (ExcelPackage package = new ExcelPackage(new FileInfo(inputFilePath)))
            {
                // Get the first worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                if (worksheet == null|| worksheet.Dimension == null)
                {
                   // Console.WriteLine("The specified worksheet does not exist.");
                    log.Error("The specified worksheet does not exist ");
                    // Handle the case where the worksheet is not found
                     return dataRecords; 
                }

                // Read data from the worksheet
                int rowCount = worksheet.Dimension.Rows;
                for (int row = 2; row <= rowCount; row++)
                {
                    // Get the state value from the current row
                    var stateVal = worksheet.Cells[row, 3].Value;
                    if (stateVal != null && stateVal.ToString() == stateFilter)
                    {
                        // Process referred count
                        var referValue = worksheet.Cells[row, 26].Value;
                        var referredWeek = worksheet.Cells[row, 6].Value;

                        string referdateOnlyString = DateTime.Parse(referredWeek.ToString()).ToString("M/d/yyyy");
                        DateOnly referdate = DateOnly.ParseExact(referdateOnlyString, "M/d/yyyy");
                       // Console.WriteLine(referdate);
                        if (referValue != null && int.TryParse(referValue.ToString(), out int referValueNumeric))
                        {
                            if (referValueNumeric != 0)
                            {

                                // Update refer count for the date
                                if (referCountByDate.ContainsKey(referdate))
                                {
                                    referCountByDate[referdate]++;
                                }
                                else
                                {
                                    referCountByDate.Add(referdate, 1);
                                }
                            }
                        }
                        // Process admitted count
                        var admitValue = worksheet.Cells[row, 27].Value;
                        if (admitValue != null && int.TryParse(admitValue.ToString(), out int admitValueNumeric))
                        {
                            if (admitValueNumeric != 0)
                            {
                                var admittedWeek = worksheet.Cells[row, 9].Value;

                                string admitdateOnlyString = DateTime.Parse(admittedWeek.ToString()).ToString("M/d/yyyy");
                                DateOnly admitdate = DateOnly.ParseExact(admitdateOnlyString, "M/d/yyyy");

                                // Update admit count for the date
                                if (admitCountByDate.ContainsKey(admitdate))
                                {
                                    admitCountByDate[admitdate]++;
                                }
                                else
                                {
                                    admitCountByDate.Add(admitdate, 1);
                                }
                            }
                        }

                    }
                }
            }

            // Create data records from the dictionaries
            foreach (var referDate in referCountByDate.Keys)
            {
                dataRecords.Add(new DataRecord
                {
                    Referred = referCountByDate[referDate],
                    Admitted = admitCountByDate.ContainsKey(referDate) ? admitCountByDate[referDate] : 0,
                    ReferredWeek = referDate
                });
            }
            log.Info("Reading excel data for the state " + stateFilter + " Completed");
            return dataRecords;
        }
    }
}
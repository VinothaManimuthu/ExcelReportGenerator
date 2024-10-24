using OfficeOpenXml;
using log4net;
using log4net.Config;
using System.Collections.Generic;
namespace ExcelReportGenerator
{
    public class GenerateExcelReport
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        /// <summary>
        /// Generates an Excel report based on the provided data records.
        /// </summary>
        /// <param name="dataRecords">The list of data records to include in the report.</param>
        /// <param name="outputFilePath">The file path where the report will be saved.</param>
        /// <param name="numOfWeeks">The number of weeks to include in the report.</param>
        public static void GenerateReport(List<DataRecord> dataRecords, string outputFilePath,int numOfWeeks)
        {
            XmlConfigurator.Configure(new FileInfo("log4net.config"));
            log.Info("Excel report generation starts here");
            // Create a new Excel package
            using (ExcelPackage package = new ExcelPackage())
            {
                // Create a new worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Report");

                // Define header row
                worksheet.Cells[1, 1].Value = "Week";
                worksheet.Cells[1, numOfWeeks + 2].Value = "Average";
                worksheet.Cells[1, numOfWeeks + 3].Value = "Admitted";
                worksheet.Cells[1, numOfWeeks + 4].Value = "Referred";
                worksheet.Cells[1, numOfWeeks + 5].Value = "ReferredDate";

                // Define column widths
                for (int i = 1; i <= numOfWeeks + 5; i++)
                {
                    worksheet.Column(i).Width = 10;
                }

                // Define header row style
                worksheet.Cells[1, 1, 1, numOfWeeks+6].Style.Font.Bold = true;
                worksheet.Cells[1, 1, 1, numOfWeeks + 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, 1, 1, numOfWeeks + 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                // Group data records by ReferredWeek for the latest 'numOfWeeks'
                var dataByWeek = dataRecords.GroupBy(record => record.ReferredWeek)
                                            .OrderByDescending(group => group.Key).Take(numOfWeeks);

                              
                int row = 2;
                int weekNumber = 0;
                int countColumn = numOfWeeks+1;
                double totalConversionRatio = 0;
                // Populate data rows
                foreach (var weekGroup in dataByWeek)
                {
                    // Calculate weekend dates
                    // DateOnly currentWeekStart = weekGroup.Key.AddDays(-(int)weekGroup.Key.DayOfWeek); // Sunday
                    // DateOnly currentWeekEnd = currentWeekStart.AddDays(6); // Saturday

                    // Add data to the worksheet
                     worksheet.Cells[row, 1].Value = $"W{weekNumber}";
                     worksheet.Cells[1, countColumn].Value = $"{weekGroup.Key.ToString("MMM-d")}";// - {currentWeekEnd.ToString("MM/dd/yyyy")}";

                    // Calculate total admitted and referred for the week
                    int totalAdmitted = weekGroup.Sum(record => record.Admitted);
                    int totalReferred = weekGroup.Sum(record => record.Referred);

                    // Write data to the worksheet
                    worksheet.Cells[row, numOfWeeks + 3].Value = totalAdmitted;
                    worksheet.Cells[row, numOfWeeks + 4].Value = totalReferred;
                    worksheet.Cells[row, numOfWeeks + 5].Value = weekGroup.Key.ToString("M/d/yyyy"); // ReferredDate

                    // Calculate conversion ratio
                    double conversionRatio = 0;
                    if (totalAdmitted + totalReferred > 0)
                    {
                        conversionRatio = totalReferred > 0 ? (double)totalAdmitted /  totalReferred * 100:0;
                        totalConversionRatio += conversionRatio;
                    }
                    worksheet.Cells[row, countColumn].Value = conversionRatio.ToString("0.00") + "%";

                    row++;
                    countColumn--;
                    weekNumber++;
                }
                //Calculate Average
                if (numOfWeeks > 0)
                {
                    double averageConversionRatio = totalConversionRatio / numOfWeeks;
                    worksheet.Cells[2, numOfWeeks + 2].Value = averageConversionRatio.ToString("0.00") + "%"; // Place average in the next row
                }

                // Save the Excel file
                package.SaveAs(new FileInfo(outputFilePath));
            }
            //Console.WriteLine("Excel report generated successfully!");
            log.Info("Excel Report generation Completed");
        }

       
    }
    
}
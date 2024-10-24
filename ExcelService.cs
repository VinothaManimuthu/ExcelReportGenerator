using System;
using OfficeOpenXml;
using log4net;
using log4net.Config;
namespace ExcelReportGenerator
{
    public class ExcelService
    {
        // private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        private readonly ILog log;
        private readonly IReadDataFromExcel _readDataFromExcel;
        private readonly IGenerateExcelReport _generateExcelReport;

        // Add constructor accepting three parameters
        public ExcelService(ILog logger, IReadDataFromExcel readDataFromExcel, IGenerateExcelReport generateExcelReport)
        {
            log = logger;
            _readDataFromExcel = readDataFromExcel;
            _generateExcelReport = generateExcelReport;
        }
        public ExcelService() // Default constructor if needed
        {
            log = LogManager.GetLogger(typeof(ExcelService)); // Create logger here if default is preferred
        }
        public void Run()
        {
           
            try
            {
                // Configure log4net logging
                XmlConfigurator.Configure(new FileInfo("log4net.config"));
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Specify the folder from which to read files

                string inputFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "InputFiles");  // Folder where files are stored
                string outputFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputFiles"); // Folder for the generated report

                // Retrieve all .xlsx files from the folder
                string[] excelFiles = Directory.GetFiles(inputFolderPath, "*.xlsx");

                // Check if any .xlsx files are found
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No .xlsx files found in the specified folder.");
                    log.Error("No .xlsx files found in the folder: " + inputFolderPath);
                    return;
                }

                // Picks .xlsx file found in the folder
                string inputFilePath = excelFiles[0];
                Console.WriteLine($"Using file: {Path.GetFileName(inputFilePath)}");

                // Get current date and time
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                // Append the timestamp to the output file name
                string outputFilePath = Path.Combine(outputFolderPath, $"GeneratedReport_{timestamp}.xlsx");
               
                string? stateFilter;
                // Loop until a valid state is provided
                while (true)
                {
                    Console.WriteLine("Enter the state to filter (e.g., MN, IA, WI):");
                    stateFilter = Console.ReadLine();

                    if (!string.IsNullOrEmpty(stateFilter))
                    {
                        if (!IsAlphabetic(stateFilter))
                        {
                            // Log and display the error for invalid input
                            Console.WriteLine("Error: State filter must contain only alphabetic characters.");
                            log.Error("State filter must contain only alphabetic characters.");
                        }
                        else
                        {
                            // Valid state input, break the loop
                            stateFilter = stateFilter.ToUpper(); // Convert state filter to uppercase for consistency
                            log.Info("User entered state " + stateFilter);
                            break;
                        }
                    }
                    else
                    {
                        // Prompt again if the state filter is empty
                        Console.WriteLine("Please provide the state data for filtering");
                        log.Error("Please provide the state data for filtering");
                    }
                }


                // Read data from input Excel file
                List<DataRecord> dataRecords = ReadDataFromExcel.ReadData(inputFilePath, stateFilter);
                // Check if dataRecords has values
                if (dataRecords == null || dataRecords.Count == 0)
                {
                    Console.WriteLine("No records found for the specified state.");
                    log.Warn($"No records found for state: {stateFilter}");
                    return; // Exit the application if no records are found
                }

                string? inputWeeks;
                int numOfWeeks;
                // Loop until a valid number of weeks is provided
                while (true)
                {
                    Console.WriteLine($"Enter the number of weeks the reports need to generate(within {dataRecords.Count}) : ");
                    inputWeeks = Console.ReadLine();

                    if (int.TryParse(inputWeeks, out int result))
                    {
                        // Check if the entered number is within the valid range
                        if (result > 0 && result <= dataRecords.Count)
                        {
                            numOfWeeks = result;
                            break; // Valid number of weeks, exit the loop
                        }
                        else
                        {
                            Console.WriteLine($"Please enter a number between 1 and {dataRecords.Count}.");
                            log.Error($"User entered a number of weeks out of range: {result}");
                        }
                    }
                    else
                    {
                        // Log and display the error for invalid input
                        Console.WriteLine("Invalid input. Please enter a valid number of weeks.");
                        log.Error("Invalid data for input week");
                    }
                }

                // Generate Excel report after valid input is received
                GenerateExcelReport.GenerateReport(dataRecords, outputFilePath, numOfWeeks);

            }
            catch (FileNotFoundException ex)
            {
                // Handle the file not found exception
                log.Error($"File not found: {ex.Message}");
                log.Error($"StackTrace: {ex.StackTrace}");
                Console.WriteLine($"Error: {ex.Message}");


            }
            catch (Exception ex)
            {
               // Console.WriteLine($"Error: {ex.Message}");
                // Handle the exception
                log.Error("Error:: " + ex.Message);
                log.Error($"StackTrace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Checks if the input string contains only alphabetic characters.
        /// </summary>
        /// <param name="input">The input string to validate.</param>
        /// <returns>True if the input is alphabetic; otherwise, false.</returns>
        public bool IsAlphabetic(string input)
        {
            foreach (char c in input)
            {
                // Check if each character is an alphabet
                if (!char.IsLetter(c))
                {
                    return false; // Return false if any character is not a letter
                }
            }
            return true; // All characters are alphabetic
        }
    }
}
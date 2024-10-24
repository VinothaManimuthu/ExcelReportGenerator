using System;
using System.Collections.Generic;
namespace ExcelReportGenerator
{
    public interface IGenerateExcelReport
    {
       void GenerateReport(List<DataRecord> dataRecords, string outputFilePath, int numOfWeeks);
    }
}

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Components.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using TopTenErrors.Models;


namespace TopTenErrors.Services
{
	public class TxtFileServices
	{
 
        public bool IsFilesExists(IFormFileCollection TopTenTxtFiles)
        {

            foreach (var file in TopTenTxtFiles)
            {
                if (file.Length == 0 || file == null)
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsValidFilesExtentions(IFormFileCollection TopTenTxtFiles)
        {
            foreach (var file in TopTenTxtFiles)
            {
                if (Path.GetExtension(file.FileName).ToLower() != ".txt")
                {
                    return false;
                }
            }

            return true;
        }

        public string TxtReader(IFormFileCollection TopTenTxtFiles)
        {
            string TxtText = "";
            foreach (var file in TopTenTxtFiles)
            {
                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    TxtText += reader.ReadToEnd();
                }
            }
            return TxtText;
        }

        public List<ExcelFile> TxtFilter(IFormFileCollection TopTenTxtFiles)
        {
            string TxtContent = TxtReader(TopTenTxtFiles);
            List<ExcelFile> excelFiles = new List<ExcelFile>();
            List<string> reportDatesList = ExtractValues(TxtContent, @"REPORT DATE [ 0-9]{2}-[ 0-9]{2}-[0-9]{2}", 11);
            List<string> programList = ExtractValues(TxtContent, @"PROGRAM: [A-Z]{2}[0-9]{4}", 8);
            List<string> orgCntrList = ExtractValues(TxtContent, @"ORIG CTR  [A-Z0-9]{1}[A-Z0-9]{1}[0-9A-Z]{1}", 8);
            List<string> weekEndingList = ExtractValues(TxtContent, @"WEEK ENDING [ 0-9]{2}-[ 0-9]{2}-[0-9]{2}", 11);
            List<string> errorNo = ExtractValues(TxtContent, @"E[0-9]{4}", 0);
            List<string> noOfError = ExtractValues(TxtContent, @"\bE\d{4}\s+\d+",5);
            List<string> errorMessage = ExtractValues(TxtContent, @"\bE\d{4}\s+\d+\s+((?:(?!\bE\d{4}|REPORT CONTINUED|END-OF-REPORT).)*)", 15); //this gets the whole body info

            reportDatesList = FixReportDate(reportDatesList);
            weekEndingList = FixReportDate(weekEndingList);

            for (int i = 0; i < errorMessage.Count; i++)
            {
                
                if (reportDatesList.Count <= i || programList.Count <= i || orgCntrList.Count <= i || weekEndingList.Count <= i)
                {
                   while (IsDifferenceOneOrMoreDays(reportDatesList[i-1], weekEndingList[i-1]) == true)
                   {
                       reportDatesList.Add(reportDatesList[i - 1]);
                       programList.Add(programList[i - 1]);
                       orgCntrList.Add(orgCntrList[i - 1]);
                       weekEndingList.Add(weekEndingList[i - 1]);
                   }
                }
                
                
                ExcelFile excelFile = new ExcelFile
                (
                    reportDatesList[i],
                    programList[i],
                    weekEndingList[i],
                    orgCntrList[i],
                    errorNo[i],
                    noOfError[i],
                    errorMessage[i],
                    Count_MonthYear(weekEndingList[i])
                );
                excelFiles.Add(excelFile);
            }
            return excelFiles;
        }

        private List<string> ExtractValues(string inputString, string pattern, int startIndex)
            {
            List<string> valuesList = new List<string>();
            MatchCollection matches = Regex.Matches(inputString, pattern);
            foreach (Match match in matches)
            {
                string value = match.Value.ToString().Substring(startIndex).Trim();
                valuesList.Add(value);
            }
            return valuesList;
        }

        private string Count_MonthYear(string inputDate)
        {
            DateTime date;
            if (DateTime.TryParse(inputDate, out date))
            {
                return date.ToString("yyyy MMMM");
            }
            else
            {
                throw new ArgumentException("Invalid date format. Expected format: M/d/yyyy");
            }
        }

        private string CreateFileName(List<ExcelFile> topTenErrorsObject)
        {
            var file = topTenErrorsObject.Select(fileName => fileName.MONTH_YEAR).FirstOrDefault();
            DateTime date;
            if (DateTime.TryParse(file, out date))
            {
                return date.ToString("yyyy MMMM" + "Top Ten Errors");
            }
            else
            {
                throw new ArgumentException("Invalid date format. Expected format: M/d/yyyy");
            }

        }

        //create a excel sheet 
        public async Task CreateExcelSheet(List<ExcelFile> topTenErrorsObject)
        {
            string fileName = CreateFileName(topTenErrorsObject);
            var WorkSheetName = topTenErrorsObject.Select(fileName => fileName.MONTH_YEAR).FirstOrDefault();
            FileInfo file = new FileInfo(fileName);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(WorkSheetName);

                //range -> representing range of cells | column 
                var range = worksheet.Cells["A1"].LoadFromCollection(topTenErrorsObject, true);
                range.AutoFitColumns();

                //styling Header
                worksheet.Cells["A1"].Style.Font.Bold = true;
                worksheet.Cells["A1"].Style.Font.Size = 14;
                worksheet.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Set the horizontal and vertical alignment of all cells
                var style = worksheet.Cells.Style;
                style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                style.VerticalAlignment = ExcelVerticalAlignment.Center;

                // Set the width of all columns to 20
                for (int i = 1; i <= worksheet.Dimension.End.Column; i++)
                {
                    worksheet.Column(i).Width = 20;
                }

                await package.SaveAsync();
            }
        }
        private List<String> FixReportDate(List<String> inputDate)
        {

            List<String> dates = new List<String>();
            foreach (string strDate in inputDate)
            {
                string[] dateParts = strDate.Split('-');
                int month = int.Parse(dateParts[0]);
                int day = int.Parse(dateParts[1]);
                int year = int.Parse("20" + dateParts[2]);
                DateTime date = new DateTime(year, month, day);
                dates.Add(date.ToString("M/d/yyyy"));
            }
            return dates;
        }
        private bool IsDifferenceOneOrMoreDays(string date1Str,string date2str)
        {
            DateTime date1 = DateTime.ParseExact(date1Str, "M/d/yyyy", CultureInfo.InvariantCulture);

            DateTime date2 = DateTime.ParseExact(date2str, "M/d/yyyy", CultureInfo.InvariantCulture);

            TimeSpan difference = date2 - date1;
            int differenceInDays = (int)difference.TotalDays;
            return differenceInDays == 1;
        }
    }
}


using OfficeOpenXml;
using ScoutnetExport.Models;
using System.Data;

namespace ScoutnetExport.Helpers
{
    public static class ExportHelper
    {
        public static byte[] ExportToMunicipalReport(
            string filePath,
            Dictionary<string, Dictionary<DateOnly, string>> occasions,
            Dictionary<string, List<Participant>> departments)
        {
            using (FileStream sourceFileStream = File.OpenRead(filePath))
            {
                using (var package = new ExcelPackage())
                {
                    package.Load(sourceFileStream);

                    package.AddDataToRegister(departments);

                    package.IncreaseWorksheetsToFitDepartmentAmount(departments, 3);

                    package.AddDataForDepartments(occasions, departments, 3);

                    return package.GetAsByteArray();
                }
            }
        }

        private static void AddDataToRegister(
            this ExcelPackage package,
            Dictionary<string,
            List<Participant>> participants)
        {
            var registerSheet = package.Workbook.Worksheets[2];
            var registerCurrentRow = 7;
            var registerId = 1;

            foreach (var department in participants)
            {
                foreach (var participantData in department.Value)
                {
                    var existingMember = participants.Where(x => x.Value.Any(x => x.MemberNumber == participantData.MemberNumber && x.Id > 0));

                    if (existingMember.Any())
                    {
                        var existing = existingMember.First().Value.FirstOrDefault(x => x.MemberNumber == participantData.MemberNumber);

                        if (existing != null)
                        {
                            participantData.Id = existing.Id;
                        }
                    }
                    else
                    {
                        registerSheet.Cells[registerCurrentRow, 1].Value = registerId;
                        registerSheet.Cells[registerCurrentRow, 2].Value = participantData.Name;
                        registerSheet.Cells[registerCurrentRow, 3].Value = participantData.BirhtDate.ToString("yyyyMMdd");
                        registerSheet.Cells[registerCurrentRow, 4].Value = participantData.Gender == "M" ? 2 : 1;

                        participantData.Id = registerId;

                        registerCurrentRow++;
                        registerId++;
                    }
                }
            }
        }

        private static void IncreaseWorksheetsToFitDepartmentAmount(
            this ExcelPackage package,
            Dictionary<string, List<Participant>> departments,
            int startWorksheet)
        {
            if (departments.Count > 3)
            {
                var sourceSheet = package.Workbook.Worksheets[startWorksheet];

                for (int i = startWorksheet + 1; i <= departments.Count; i++)
                {
                    var originalName = sourceSheet.Name.Split(" ");

                    package.Workbook.Worksheets.Copy(sourceSheet.Name, $"{originalName[0]} {i}");
                }
            }
        }

        private static void AddDataForDepartments(
            this ExcelPackage package,
            Dictionary<string, Dictionary<DateOnly, string>> dates,
            Dictionary<string, List<Participant>> departments,
            int startWorksheet)
        {
            var worksheetNumber = startWorksheet;

            foreach (var department in departments)
            {
                var worksheet = package.Workbook.Worksheets[worksheetNumber];

                worksheet.AddOccasions(dates, department.Key);
                worksheet.AddDataForChildren(department.Value);
                worksheet.AddDataForAdults(department.Value);

                worksheetNumber++;
            }
        }

        private static void AddOccasions(
            this ExcelWorksheet worksheet,
            Dictionary<string, Dictionary<DateOnly, string>> occasions,
            string department)
        {
            var dateColumn = 97;

            foreach (var occasion in occasions[department].OrderBy(date => date.Key))
            {
                worksheet.Cells[6, dateColumn].Value = occasion.Value;
                worksheet.Cells[16, dateColumn].Value = occasion.Key.Month;
                worksheet.Cells[17, dateColumn].Value = occasion.Key.Day;
                dateColumn++;
            }
        }

        private static void AddDataForChildren(
            this ExcelWorksheet worksheet,
            List<Participant> participants)
        {
            var children = participants.Where(x => x.Role != "Ledare").ToList();
            var currentRow = 19;

            foreach (var child in children)
            {
                if (currentRow == 38)
                {
                    currentRow = 71;
                }
                
                worksheet.Cells[currentRow, 2].Value = child.Id;
                worksheet.Cells[currentRow, 3].Value = child.Name;

                var dateColumn = 97;

                foreach (var participationDate in child.ParticipationDates.OrderBy(entry => entry.Key))
                {
                    if (participationDate.Value)
                    {
                        worksheet.Cells[currentRow, dateColumn].Value = "1";
                    }

                    dateColumn++;
                }

                currentRow++;
            }
        }

        private static void AddDataForAdults(
            this ExcelWorksheet worksheet,
            List<Participant> participants)
        {
            var adults = participants.Where(x => x.Role == "Ledare").ToList();
            var currentRow = 38;

            foreach (var adult in adults)
            {
                if (adults.IndexOf(adult) == 2)
                {
                    currentRow = 110;
                }

                worksheet.Cells[currentRow, 2].Value = adult.Id;
                worksheet.Cells[currentRow, 4].Value = adult.Name;

                var dateColumn = 97;

                foreach (var participationDate in adult.ParticipationDates.OrderBy(entry => entry.Key))
                {
                    if (participationDate.Value)
                    {
                        worksheet.Cells[currentRow, dateColumn].Value = "1";
                    }

                    dateColumn++;
                }

                currentRow++;
            }
        }
    }
}

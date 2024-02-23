using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using ScoutnetExport.Models;

namespace ScoutnetExport.Helpers
{
    public static class ImportHelper
    {
        public static (Dictionary<string, Dictionary<DateOnly, string>>, Dictionary<string, List<Participant>>) ImportDataFromExcelReport(Stream stream)
        {
            var occasions = new Dictionary<string, Dictionary<DateOnly, string>>();
            var departments = new Dictionary<string, List<Participant>>();

            using (IWorkbook book1 = new HSSFWorkbook(stream))
            {
                for (int sheetNumber = 0; sheetNumber < book1.NumberOfSheets; sheetNumber++)
                {
                    var inputBook = book1.GetSheetAt(sheetNumber);
                    var currentRowNum = 2;
                    var labelsRow = inputBook.GetRow(0);
                    var datesRow = inputBook.GetRow(1);

                    var dates = new Dictionary<DateOnly, string>();

                    foreach (var cell in datesRow.Cells)
                    {
                        if (DateOnly.TryParse(cell.StringCellValue, out DateOnly dateValue) && dateValue > DateOnly.MinValue)
                        {
                            dates.Add(dateValue, labelsRow.Cells[cell.ColumnIndex].StringCellValue);
                        }
                    }

                    occasions.Add(inputBook.SheetName, dates);

                    if (inputBook.LastRowNum > 1)
                    {
                        var participantsList = new List<Participant>();

                        while (currentRowNum < inputBook.LastRowNum)
                        {
                            var currentRow = inputBook.GetRow(currentRowNum);
                            Int32.TryParse(currentRow.Cells[4].NumericCellValue.ToString(), out int zipcode);

                            if (DateOnly.TryParse(currentRow.Cells[5].StringCellValue, out DateOnly birthDate))
                            {
                                var person = new Participant()
                                {
                                    Groupname = inputBook.SheetName,
                                    PatrolName = currentRow.Cells[0].StringCellValue,
                                    MemberNumber = currentRow.Cells[1].NumericCellValue.ToString(),
                                    Name = currentRow.Cells[2].StringCellValue,
                                    Gender = currentRow.Cells[3].StringCellValue,
                                    BirhtDate = birthDate,
                                    ZipCode = zipcode,
                                    Role = currentRow.Cells.LastOrDefault()?.StringCellValue
                                };

                                for (int i = 6; i < currentRow.Cells.Count - 1; i++)
                                {
                                    DateOnly.TryParse(datesRow.Cells[i].StringCellValue, out DateOnly dateValue);

                                    try
                                    {
                                        if (currentRow.Cells[i].NumericCellValue == 1)
                                        {
                                            person.ParticipationDates.Add(dateValue, true);
                                        }
                                        else
                                        {
                                            person.ParticipationDates.Add(dateValue, false);
                                        }
                                    }
                                    catch
                                    {
                                        person.ParticipationDates.Add(dateValue, false);
                                    }
                                }

                                participantsList.Add(person);
                            }

                            currentRowNum++;
                        }

                        departments.Add(inputBook.SheetName, participantsList);
                    }
                }
            }

            return (occasions, departments);
        }
    }
}

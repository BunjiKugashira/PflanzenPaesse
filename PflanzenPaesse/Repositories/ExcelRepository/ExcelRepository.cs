namespace PflanzenPaesse.Repositories.ExcelRepository
{
    using System.Collections.Generic;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class ExcelRepository
    {
        public static readonly string[] AllowedFileEndings = { "xlsx" };
        public const int MinRowIndex = 2;
        public const int MaxRowIndex = 1048576;

        public static IEnumerable<IDictionary<string, string>> Import(string fileName, string tableName)
        {
            using var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);
            var workBook = spreadsheetDocument.WorkbookPart;
            var workSheet = workBook.WorksheetParts.First();
            var sheetData = workSheet.Worksheet.Elements<SheetData>().First();

            IEnumerable<string> headers = null;
            foreach (var row in sheetData.Elements<Row>())
            {
                var cells = row.Elements<Cell>();
                if (headers == null)
                {
                    headers = cells.Select(cell => cell.CellValue.Text);
                }
                else
                {
                    yield return cells.Select((cell, index) => new KeyValuePair<string, string>(headers.ElementAt(index), cell.CellValue.Text))
                        .ToDictionary(kvPair => kvPair.Key, kvPair => kvPair.Value);
                }
            }
        }

        public static int HighestUsedRowNumber(string fileName, string tableName)
        {
            using var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);
            var workBook = spreadsheetDocument.WorkbookPart;
            var workSheet = workBook.WorksheetParts.First();
            var sheetData = workSheet.Worksheet.Elements<SheetData>().First();

            return sheetData.Elements<Row>().Count();
        }
    }
}

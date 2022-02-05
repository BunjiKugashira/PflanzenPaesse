namespace PflanzenPaesse.Repositories.ExcelRepository
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Spreadsheet;

    public static class ExcelRepository
    {
        public static readonly string[] AllowedFileEndings = { "xlsx" };
        public const int MinRowIndex = 2;
        public const int MaxRowIndex = 1048576;

        private static readonly Regex CellReferenceRegex = new Regex($"^(?'column'[A-Z]+)(?'row'[0-9]+)$", RegexOptions.Compiled);

        public static IEnumerable<IDictionary<string, string>> Import(string fileName, string tableName)
        {
            using var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false);
            var workBook = spreadsheetDocument.WorkbookPart;
            var workSheet = GetWorkSheets(workBook)[tableName];
            var sheetData = workSheet.Worksheet.Elements<SheetData>().Single();
            var sharedStrings = GetSharedStrings(workBook.SharedStringTablePart).ToArray();

            IDictionary<string, string> headers = null;
            foreach (var row in sheetData.Elements<Row>())
            {
                var cells = row.Elements<Cell>();
                Console.WriteLine(string.Join(" ", cells.Select(cell => cell.CellValue?.InnerText)));
                if (headers == null)
                {
                    headers = cells.ToDictionary(cell => cell.CellReference.Value, cell => GetCellValueAsString(cell, sharedStrings));
                }
                else
                {
                    yield return cells.Select(cell => new KeyValuePair<string, string>(GetHeader(headers, cell.CellReference.Value), GetCellValueAsString(cell, sharedStrings)))
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

        private static string GetHeader(IDictionary<string, string> headers, string cellReference)
        {
            if (cellReference == null)
            {
                return null;
            }

            var column = CellReferenceRegex.Match(cellReference).Groups["column"].Value;
            var row = CellReferenceRegex.Match(headers.First().Key).Groups["row"].Value;
            var headerReference = column + row;

            return headers.TryGetValue(headerReference, out var header) ? header : string.Empty;
        }

        private static IEnumerable<string> GetSharedStrings(SharedStringTablePart sharedStringTablePart)
        {
            using var reader = OpenXmlReader.Create(sharedStringTablePart);
            while (reader.Read())
            {
                if (reader.ElementType == typeof(SharedStringItem))
                {
                    var ssi = (SharedStringItem)reader.LoadCurrentElement();
                    yield return ssi.Text.Text;
                }
            }
        }

        private static string GetCellValueAsString(Cell cell, string[] sharedStrings)
        {
            var text = cell.CellValue?.Text ?? string.Empty;

            if (cell.DataType == null)
            {
                return text;
            }

            if (cell.DataType == CellValues.SharedString)
            {
                return sharedStrings[Convert.ToInt32(text)];
            }

            return text;
        }

        private static IDictionary<string, WorksheetPart> GetWorkSheets(WorkbookPart workbook)
        {
            return workbook.Workbook.Sheets.Elements<Sheet>().ToDictionary(sheet => sheet.Name.Value, sheet => (WorksheetPart)workbook.GetPartById(sheet.Id));
        }
    }
}

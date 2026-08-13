using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

public class MonteCarloLogicTests
{
    [Fact]
    public void Percentiles_Are_Calculated_For_The_Result_Set()
    {
        var values = new[] { 10.0, 20.0, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0, 90.0, 100.0 };

        var summary = MonteCarloCalculator.GetSummary(values);

        Assert.Equal(50.0, summary.P50, 1);
        Assert.Equal(80.0, summary.P80, 1);
        Assert.Equal(20.0, summary.P20, 1);
    }

    [Fact]
    public void Triangular_Samples_Stay_Within_Min_Max_Range()
    {
        var random = new Random(42);

        for (int i = 0; i < 1000; i++)
        {
            var value = DistributionSampler.SampleTriangular(10, 25, 60, random);
            Assert.InRange(value, 10.0, 60.0);
        }
    }

    [Fact]
    public void SpreadsheetReader_Reads_Unordered_Excel_Columns()
    {
        var filePath = Path.Combine(Path.GetTempPath(), $"montecarlo-{Guid.NewGuid():N}.xlsx");

        try
        {
            CreateWorkbook(filePath, new[]
            {
                new[] { "Distribution", "Most Likely", "Max Value", "Name", "Min Value" },
                new[] { "triangular", "25", "60", "Task A", "10" },
                new[] { "uniform", "7", "10", "Task B", "1" }
            });

            var rows = SpreadsheetReader.ReadInputRows(filePath);

            Assert.Equal(2, rows.Count);
            Assert.Equal("Task A", rows[0].Name);
            Assert.Equal(10.0, rows[0].Min);
            Assert.Equal(25.0, rows[0].MostLikely);
            Assert.Equal(60.0, rows[0].Max);
            Assert.Equal("triangular", rows[0].Distribution);
        }
        finally
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }

    private static void CreateWorkbook(string filePath, string[][] rows)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = spreadsheetDocument.WorkbookPart!.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" };
        sheets.Append(sheet);

        var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
        sharedStringTablePart.SharedStringTable = new SharedStringTable();

        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        foreach (var rowValues in rows)
        {
            var row = new Row();
            for (var columnIndex = 0; columnIndex < rowValues.Length; columnIndex++)
            {
                var value = rowValues[columnIndex];
                var cell = new Cell
                {
                    CellReference = GetCellReference(columnIndex + 1, sheetData!.ChildElements.Count + 1),
                    DataType = CellValues.SharedString
                };

                var sharedStringIndex = AddSharedStringItem(sharedStringTablePart, value);
                cell.CellValue = new CellValue(sharedStringIndex.ToString());
                row.Append(cell);
            }

            sheetData.Append(row);
        }

        workbookPart.Workbook.Save();
    }

    private static int AddSharedStringItem(SharedStringTablePart sharedStringTablePart, string value)
    {
        var sharedStrings = sharedStringTablePart.SharedStringTable;
        var item = new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(value));
        sharedStrings.Append(item);
        return sharedStrings.Count() - 1;
    }

    private static string GetCellReference(int column, int row)
    {
        var letters = new StringBuilder();
        var current = column;
        while (current > 0)
        {
            var remainder = (current - 1) % 26;
            letters.Insert(0, (char)('A' + remainder));
            current = (current - 1) / 26;
        }

        return $"{letters}{row}";
    }
}

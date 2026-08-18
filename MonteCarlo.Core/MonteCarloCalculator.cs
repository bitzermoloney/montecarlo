using System.Globalization;
using System.Text;

namespace MonteCarlo.Core;

public sealed record InputRow(string Name, double Min, double MostLikely, double Max, string Distribution);
public sealed record Summary(double P20, double P50, double P80);
public sealed record SimulationResult(string Name, int Iteration, double Value);
public sealed record SummaryRow(string Name, double P20, double P50, double P80);

public static class MonteCarloCalculator
{
    public static List<double> RunSimulation(InputRow row, int iterations, Random? random = null) // Creates a list of random samples generated
    {
        var generator = random ?? new Random();
        var results = new List<double>(iterations);

        for (var i = 0; i < iterations; i++)
        {
            results.Add(DistributionSampler.Sample(row, generator));
        }

        return results;
    }

    public static Summary GetSummary(IEnumerable<double> values) // Returns the P20, P50 and P80 of each value in the list of random samples
    {
        var sorted = values.OrderBy(v => v).ToArray();

        if (sorted.Length == 0)
        {
            return new Summary(0, 0, 0);
        }

        return new Summary(
            GetPercentile(sorted, 0.20),
            GetPercentile(sorted, 0.50),
            GetPercentile(sorted, 0.80));
    }

    private static double GetPercentile(IReadOnlyList<double> sortedValues, double percentile)
    { // Gets the percentile (para) of the sorted list of values (para), returns (function)
        if (sortedValues.Count == 1)
        {
            return sortedValues[0];
        }

        var index = percentile * (sortedValues.Count - 1);
        var lowerIndex = (int)Math.Floor(index);
        var upperIndex = (int)Math.Ceiling(index);

        if (lowerIndex == upperIndex)
        {
            return sortedValues[lowerIndex];
        }

        var fraction = index - lowerIndex;
        return sortedValues[lowerIndex] + (sortedValues[upperIndex] - sortedValues[lowerIndex]) * fraction;

        // returns the value at the calculated percentile index, interpolating between the two nearest values if necessary
    }
}

public static class DistributionSampler
{
    public static double SampleTriangular(double min, double mode, double max, Random random)
    { // Finds max and min values for triangular
    // If min is greater than max, swap them
        if (min > max)
        {
            (min, max) = (max, min);
        }

        if (mode < min || mode > max)
        {
            mode = Math.Clamp(mode, min, max);
        }

        var u = random.NextDouble();

        if (Math.Abs(max - min) < double.Epsilon)
        {
            return min;
        }

        var c = (mode - min) / (max - min);

        if (u <= c)
        {
            return min + Math.Sqrt(u * (max - min) * (mode - min));
        }

        return max - Math.Sqrt((1.0 - u) * (max - min) * (max - mode));
    }

    public static double SampleUniform(double min, double max, Random random)
    { //Uniform
        return min + (max - min) * random.NextDouble();
    }

    public static double SampleNormal(double min, double mode, double max, Random random)
    { // Normal
        var sigma = (max - min) / 6.0;
        sigma = sigma <= 0 ? 1 : sigma;

        double sample;
        using var randomNumberGenerator = System.Security.Cryptography.RandomNumberGenerator.Create();
        var bytes = new byte[8];
        randomNumberGenerator.GetBytes(bytes);
        var u1 = BitConverter.ToUInt64(bytes, 0) / (double)ulong.MaxValue;
        var u2 = random.NextDouble();
        var z0 = Math.Sqrt(-2.0 * Math.Log(u1 == 0 ? 0.0000001 : u1)) * Math.Cos(2.0 * Math.PI * u2);
        sample = mode + z0 * sigma;

        if (sample < min)
        {
            sample = min;
        }

        if (sample > max)
        {
            sample = max;
        }

        return sample;
    }

    public static double Sample(InputRow row, Random random)
    { // Samples a value based on the distribution type specified in the InputRow
        var distribution = (row.Distribution ?? string.Empty).Trim();

        return distribution.ToLowerInvariant() switch
        {
            "triangular" or "tri" or "triangle" => SampleTriangular(row.Min, row.MostLikely, row.Max, random),
            "uniform" or "u" => SampleUniform(row.Min, row.Max, random),
            "normal" or "gaussian" or "n" => SampleNormal(row.Min, row.MostLikely, row.Max, random),
            _ => SampleTriangular(row.Min, row.MostLikely, row.Max, random)
        };
    }
}

public static class SpreadsheetParser // yawn
{
    public static List<InputRow> ParseCsv(string content)
    {
        var lines = content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n', StringSplitOptions.RemoveEmptyEntries);

        if (lines.Length < 2)
        {
            return new List<InputRow>();
        }

        var headers = ParseCsvLine(lines[0]);
        var rows = new List<InputRow>();

        for (var i = 1; i < lines.Length; i++)
        {
            var values = ParseCsvLine(lines[i]);
            var row = BuildFromHeaderValues(headers, values, i + 1);
            if (row is not null)
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    public static List<InputRow> ParseExcel(Stream stream)
    {
        using var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("Workbook is invalid.");
        var sheet = workbookPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().FirstOrDefault() ?? throw new InvalidOperationException("Workbook contains no worksheet.");
        var worksheetPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var rows = new List<InputRow>();
        var sheetData = worksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().ToList();

        if (sheetData.Count == 0)
        {
            return rows;
        }

        var headerCells = sheetData[0].Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToList();
        var headers = headerCells.Select(cell => GetCellValue(cell, workbookPart.SharedStringTablePart)).ToArray();

        for (var rowIndex = 1; rowIndex < sheetData.Count; rowIndex++)
        {
            var cells = sheetData[rowIndex].Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToList();
            var values = cells.Select(cell => GetCellValue(cell, workbookPart.SharedStringTablePart)).ToArray();
            var row = BuildFromHeaderValues(headers, values, rowIndex + 1);
            if (row is not null)
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    private static InputRow? BuildFromHeaderValues(string[] headers, string[] values, int rowNumber)
    {
        var dictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < headers.Length && index < values.Length; index++)
        {
            var header = NormalizeHeader(headers[index]);
            if (!string.IsNullOrWhiteSpace(header))
            {
                dictionary[header] = values[index];
            }
        }

        if (!TryGetValue(dictionary, new[] { "min", "minimum", "min value", "low", "lower bound" }, out var minText))
        {
            return null;
        }

        if (!TryGetValue(dictionary, new[] { "most likely", "mostlikely", "likely", "mode", "peak" }, out var mostLikelyText))
        {
            return null;
        }

        if (!TryGetValue(dictionary, new[] { "max", "maximum", "max value", "high", "upper bound" }, out var maxText))
        {
            return null;
        }

        var distribution = TryGetValue(dictionary, new[] { "distribution", "dist", "type" }, out var distText) ? distText : "triangular";
        var name = TryGetValue(dictionary, new[] { "name", "item", "variable", "row name" }, out var nameText) ? nameText : $"Row {rowNumber}";

        return new InputRow(
            name,
            ParseDouble(minText),
            ParseDouble(mostLikelyText),
            ParseDouble(maxText),
            distribution);
    }

    private static bool TryGetValue(IReadOnlyDictionary<string, string> values, IEnumerable<string> possibleKeys, out string value)
    {
        foreach (var key in possibleKeys)
        {
            if (values.TryGetValue(NormalizeHeader(key), out value!))
            {
                return true;
            }
        }

        value = string.Empty;
        return false;
    }

    private static string NormalizeHeader(string input)
    {
        return input
            .Replace("_", " ")
            .Replace("-", " ")
            .Replace("/", " ")
            .Trim()
            .ToLowerInvariant();
    }

    private static string[] ParseCsvLine(string line)
    {
        var values = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var character = line[i];
            if (character == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (character == ',' && !inQuotes)
            {
                values.Add(current.ToString());
                current.Clear();
            }
            else
            {
                current.Append(character);
            }
        }

        values.Add(current.ToString());
        return values.ToArray();
    }

    private static double ParseDouble(string text)
    {
        var candidate = text.Trim();
        if (string.IsNullOrWhiteSpace(candidate))
        {
            throw new InvalidOperationException("A numeric value could not be read from the spreadsheet.");
        }

        if (double.TryParse(candidate, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return value;
        }

        if (double.TryParse(candidate, NumberStyles.Float, CultureInfo.GetCultureInfo("en-US"), out value))
        {
            return value;
        }

        throw new InvalidOperationException($"The value '{candidate}' is not a valid number.");
    }

    private static string GetCellValue(DocumentFormat.OpenXml.Spreadsheet.Cell cell, DocumentFormat.OpenXml.Packaging.SharedStringTablePart? sharedStringTablePart = null)
    {
        if (cell.DataType is not null && cell.DataType.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
        {
            var indexText = cell.CellValue?.InnerText;
            if (int.TryParse(indexText, out var index) && sharedStringTablePart is not null)
            {
                var sharedStrings = sharedStringTablePart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>();
                if (index >= 0 && index < sharedStrings.Count())
                {
                    return sharedStrings.ElementAt(index).InnerText.Trim();
                }
            }
        }

        return cell.CellValue?.InnerText ?? string.Empty;
    }
}

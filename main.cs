using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length == 0 || args[0].Equals("--help", StringComparison.OrdinalIgnoreCase) || args[0].Equals("-h", StringComparison.OrdinalIgnoreCase))
        {
            PrintUsage();
            return 0;
        }

        try
        {
            var filePath = args[0];
            var iterations = args.Length > 1 && int.TryParse(args[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedIterations)
                ? Math.Max(1, parsedIterations)
                : 1000;

            var rows = SpreadsheetReader.ReadInputRows(filePath);

            if (rows.Count == 0)
            {
                Console.Error.WriteLine("No valid rows were found in the input spreadsheet.");
                return 1;
            }

            var allResults = new List<SimulationResult>();
            var summaries = new List<SummaryRow>();

            foreach (var row in rows)
            {
                var values = MonteCarloCalculator.RunSimulation(row, iterations);
                var summary = MonteCarloCalculator.GetSummary(values);

                allResults.AddRange(values.Select((value, index) => new SimulationResult(row.Name, index + 1, value)));
                summaries.Add(new SummaryRow(row.Name, summary.P20, summary.P50, summary.P80));
            }

            Console.WriteLine("Monte Carlo Simulation Results");
            Console.WriteLine($"Iterations per row: {iterations}");
            Console.WriteLine();

            Console.WriteLine("Summary");
            Console.WriteLine("Name | P20 | P50 | P80");
            foreach (var summary in summaries)
            {
                Console.WriteLine($"{summary.Name} | {summary.P20:F2} | {summary.P50:F2} | {summary.P80:F2}");
            }

            Console.WriteLine();
            Console.WriteLine("All simulation values");
            Console.WriteLine("Name | Iteration | Value");
            foreach (var result in allResults.Take(20))
            {
                Console.WriteLine($"{result.Name} | {result.Iteration} | {result.Value:F2}");
            }

            if (allResults.Count > 20)
            {
                Console.WriteLine("... only the first 20 rows are displayed; full results were written to simulation-results.csv");
            }

            var resultCsvPath = Path.Combine(AppContext.BaseDirectory, "simulation-results.csv");
            var summaryCsvPath = Path.Combine(AppContext.BaseDirectory, "summary.csv");

            WriteCsv(resultCsvPath, allResults, new[] { "Name", "Iteration", "Value" }, result =>
            {
                return new[] { result.Name, result.Iteration.ToString(CultureInfo.InvariantCulture), result.Value.ToString(CultureInfo.InvariantCulture) };
            });

            WriteCsv(summaryCsvPath, summaries, new[] { "Name", "P20", "P50", "P80" }, summary =>
            {
                return new[]
                {
                    summary.Name,
                    summary.P20.ToString(CultureInfo.InvariantCulture),
                    summary.P50.ToString(CultureInfo.InvariantCulture),
                    summary.P80.ToString(CultureInfo.InvariantCulture)
                };
            });

            Console.WriteLine();
            Console.WriteLine($"Output written to: {resultCsvPath}");
            Console.WriteLine($"Output written to: {summaryCsvPath}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static void PrintUsage()
    {
        Console.WriteLine("Usage: montecarlo <input.xlsx> [iterations]");
        Console.WriteLine();
        Console.WriteLine("Required columns: min value, most likely, max value, distribution");
        Console.WriteLine("Headers can appear in any order. Supported distributions: triangular, uniform, normal.");
        Console.WriteLine("Example: montecarlo input.xlsx 1000");
    }

    private static void WriteCsv<T>(string filePath, IEnumerable<T> values, string[] headers, Func<T, string[]> rowFactory)
    {
        var lines = new List<string>
        {
            string.Join(",", headers.Select(EscapeCsv))
        };

        foreach (var value in values)
        {
            lines.Add(string.Join(",", rowFactory(value).Select(EscapeCsv)));
        }

        File.WriteAllLines(filePath, lines, Encoding.UTF8);
    }

    private static string EscapeCsv(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return value;
    }
}

public sealed record InputRow(string Name, double Min, double MostLikely, double Max, string Distribution);
public sealed record SimulationResult(string Name, int Iteration, double Value);
public sealed record SummaryRow(string Name, double P20, double P50, double P80);
public sealed record Summary(double P20, double P50, double P80);

public static class MonteCarloCalculator
{
    public static List<double> RunSimulation(InputRow row, int iterations, Random? random = null)
    {
        var generator = random ?? new Random();
        var results = new List<double>(iterations);

        for (var i = 0; i < iterations; i++)
        {
            results.Add(DistributionSampler.Sample(row, generator));
        }

        return results;
    }

    public static Summary GetSummary(IEnumerable<double> values)
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
    {
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
    }
}

public static class DistributionSampler
{
    public static double SampleTriangular(double min, double mode, double max, Random random)
    {
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
    {
        return min + (max - min) * random.NextDouble();
    }

    public static double SampleNormal(double min, double mode, double max, Random random)
    {
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
    {
        var distribution = (row.Distribution ?? string.Empty).Trim();

        return distribution.ToLowerInvariant() switch
        {
            "triangular" or "tri" or "triangle" => SampleTriangular(row.Min, row.MostLikely, row.Max, random),
            "uniform" or "u" => SampleUniform(row.Min, row.Max, random),
            "normal" or "gaussian" or "n" => SampleNormal(row.Min, row.MostLikely, row.Max, random),
            _ => SampleTriangular(row.Min, row.MostLikely, row.Max, random)
        };
    }

    public static double SampleTriangular( double min, double likely, double max) => SampleTriangular(min, likely, max, new Random());
}

public static class SpreadsheetReader
{
    public static List<InputRow> ReadInputRows(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("An input file path is required.", nameof(filePath));
        }

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Input file not found: {filePath}", filePath);
        }

        var extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension == ".csv")
        {
            return ReadCsv(filePath);
        }

        if (extension == ".xlsx" || extension == ".xlsm")
        {
            return ReadXlsx(filePath);
        }

        throw new NotSupportedException("Only CSV and Excel (.xlsx/.xlsm) files are supported.");
    }

    private static List<InputRow> ReadCsv(string filePath)
    {
        var rows = new List<InputRow>();
        var lines = File.ReadAllLines(filePath);

        if (lines.Length < 2)
        {
            return rows;
        }

        var headers = ParseCsvLine(lines[0]);
        var mappedHeaders = BuildHeaderAliases(headers);

        for (var i = 1; i < lines.Length; i++)
        {
            var values = ParseCsvLine(lines[i]);
            var row = BuildInputRow(headers, values, mappedHeaders, i + 1);
            if (row is not null)
            {
                rows.Add(row);
            }
        }

        return rows;
    }

    private static List<InputRow> ReadXlsx(string filePath)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The workbook is missing a workbook part.");
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault() ?? throw new InvalidOperationException("The workbook does not contain a worksheet.");
        var part = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var rows = new List<InputRow>();
        var sheetData = part.Worksheet.Descendants<Row>().ToList();

        if (sheetData.Count == 0)
        {
            return rows;
        }

        var headerCells = sheetData[0].Descendants<Cell>().ToList();
        var headerMap = BuildCellHeaderMap(headerCells, workbookPart.SharedStringTablePart);

        for (var rowIndex = 1; rowIndex < sheetData.Count; rowIndex++)
        {
            var rowData = sheetData[rowIndex];
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var cell in rowData.Descendants<Cell>())
            {
                var columnIndex = GetColumnIndex(cell.CellReference);
                if (!headerMap.TryGetValue(columnIndex, out var headerName))
                {
                    continue;
                }

                values[headerName] = GetCellValue(cell, workbookPart.SharedStringTablePart);
            }

            var inputRow = BuildInputRow(values, rowIndex + 1);
            if (inputRow is not null)
            {
                rows.Add(inputRow);
            }
        }

        return rows;
    }

    private static InputRow? BuildInputRow(IReadOnlyDictionary<string, string> values, int rowNumber)
    {
        if (!TryGetValue(values, new[] { "min", "minimum", "min value", "low", "lower bound" }, out var minText))
        {
            return null;
        }

        if (!TryGetValue(values, new[] { "most likely", "mostlikely", "likely", "mode", "mode value", "peak" }, out var likelyText))
        {
            return null;
        }

        if (!TryGetValue(values, new[] { "max", "maximum", "max value", "high", "upper bound" }, out var maxText))
        {
            return null;
        }

        if (!TryGetValue(values, new[] { "distribution", "dist", "type" }, out var distributionText))
        {
            distributionText = "triangular";
        }

        var name = TryGetValue(values, new[] { "name", "item", "variable", "row name" }, out var nameText)
            ? nameText
            : $"Row {rowNumber}";

        var min = ParseDouble(minText);
        var likely = ParseDouble(likelyText);
        var max = ParseDouble(maxText);

        return new InputRow(name, min, likely, max, distributionText);
    }

    private static InputRow? BuildInputRow(string[] headers, string[] values, Dictionary<string, string> headerAliases, int rowNumber)
    {
        var dictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < headers.Length && index < values.Length; index++)
        {
            var normalizedHeader = NormalizeHeader(headers[index]);
            if (string.IsNullOrWhiteSpace(normalizedHeader))
            {
                continue;
            }

            dictionary[normalizedHeader] = values[index];
        }

        return BuildInputRow(dictionary, rowNumber);
    }

    private static Dictionary<int, string> BuildCellHeaderMap(IEnumerable<Cell> cells, SharedStringTablePart? sharedStringTablePart)
    {
        var map = new Dictionary<int, string>();
        foreach (var cell in cells)
        {
            var columnIndex = GetColumnIndex(cell.CellReference);
            var value = GetCellValue(cell, sharedStringTablePart);
            if (!string.IsNullOrWhiteSpace(value))
            {
                map[columnIndex] = NormalizeHeader(value);
            }
        }

        return map;
    }

    private static Dictionary<string, string> BuildHeaderAliases(string[] headers)
    {
        return headers
            .Where(header => !string.IsNullOrWhiteSpace(header))
            .ToDictionary(header => NormalizeHeader(header), header => header, StringComparer.OrdinalIgnoreCase);
    }

    private static bool TryGetValue(IReadOnlyDictionary<string, string> values, IEnumerable<string> possibleKeys, out string value)
    {
        foreach (var key in possibleKeys)
        {
            if (values.TryGetValue(key, out value!))
            {
                return true;
            }

            foreach (var pair in values)
            {
                if (NormalizeHeader(pair.Key) == NormalizeHeader(key))
                {
                    value = pair.Value;
                    return true;
                }
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

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return -1;
        }

        var letters = new StringBuilder();
        foreach (var character in cellReference)
        {
            if (char.IsLetter(character))
            {
                letters.Append(character);
            }
            else
            {
                break;
            }
        }

        var value = 0;
        foreach (var letter in letters.ToString())
        {
            value = value * 26 + (char.ToUpperInvariant(letter) - 'A' + 1);
        }

        return value;
    }

    private static string GetCellValue(Cell cell, SharedStringTablePart? sharedStringTablePart = null)
    {
        if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
        {
            var sharedStringValue = cell.CellValue?.InnerText;
            if (int.TryParse(sharedStringValue, out var sharedIndex) && sharedStringTablePart is not null)
            {
                var sharedStringTable = sharedStringTablePart.SharedStringTable;
                if (sharedIndex >= 0 && sharedIndex < sharedStringTable.Elements<SharedStringItem>().Count())
                {
                    return sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedIndex).InnerText.Trim();
                }
            }
        }

        if (cell.DataType is null)
        {
            return cell.InnerText.Trim();
        }

        if (cell.DataType == CellValues.Boolean)
        {
            return cell.CellValue?.InnerText ?? string.Empty;
        }

        return cell.CellValue?.InnerText ?? string.Empty;
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
}

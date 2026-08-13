# Monte Carlo Spreadsheet Processor

This C# console app reads an Excel spreadsheet or CSV file, detects columns like minimum, most likely, maximum, and distribution even when their order varies, runs a Monte Carlo simulation 1000 times by default, and writes a summary plus all generated values.

## Usage

```bash
dotnet run --project montecarlo.csproj input.xlsx 1000
```

The input rows must include the following concepts, with header names that can vary:

- min value / minimum / low
- most likely / mode / peak
- max value / maximum / high
- distribution (triangular, uniform, or normal)

It writes:

- a console summary with P20, P50, and P80
- a generated CSV file containing all simulation values
- a generated CSV file containing the percentiles

## Example CSV format

```csv
Name,Min Value,Most Likely,Max Value,Distribution
Task A,10,25,60,triangular
Task B,1,5,10,uniform
```

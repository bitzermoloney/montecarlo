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

## GitHub Pages and backend API setup

GitHub Pages only serves static files. It cannot run the ASP.NET upload API from this repository, so the browser must call a real backend endpoint.

Use one of these options:

- Run the API locally:
  `dotnet run --project MonteCarlo.Web`
- Deploy the ASP.NET app to Azure, Render, Fly.io, Railway, or another host and keep the `/api/upload` route there.

When the page is hosted on GitHub Pages, set the API URL before the upload runs:

```html
<script>
  window.MC_API_URL = 'https://your-backend.example.com/api/upload';
</script>
```

If you are running the app from the local ASP.NET server, the page will automatically use the same origin for `/api/upload`.

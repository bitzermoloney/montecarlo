using System.Text.Json;
using MonteCarlo.Core;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
            .AllowAnyMethod()
            .AllowAnyHeader();
    });
});

builder.Services.AddEndpointsApiExplorer();

var app = builder.Build();
app.UseCors("AllowAll");
app.UseDefaultFiles();
app.UseStaticFiles();

app.MapGet("/api/health", () => Results.Ok(new { status = "ok" }));

app.MapPost("/api/upload", async (HttpRequest request) =>
{
    if (!request.HasFormContentType)
    {
        return Results.BadRequest(new { error = "Expected multipart/form-data upload." });
    }

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null || file.Length == 0)
    {
        return Results.BadRequest(new { error = "No file was uploaded." });
    }

    var iterations = 1000;
    if (form.TryGetValue("iterations", out var iterationValue) && int.TryParse(iterationValue.FirstOrDefault(), out var parsedIterations))
    {
        iterations = Math.Max(1, parsedIterations);
    }

    using var stream = file.OpenReadStream();

    List<InputRow> rows;
    if (file.ContentType.Contains("excel", StringComparison.OrdinalIgnoreCase) || file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || file.FileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) || file.FileName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
    {
        rows = SpreadsheetParser.ParseExcel(stream);
    }
    else
    {
        using var reader = new StreamReader(stream);
        var csvText = await reader.ReadToEndAsync();
        rows = SpreadsheetParser.ParseCsv(csvText);
    }

    if (rows.Count == 0)
    {
        return Results.BadRequest(new { error = "No valid rows were found in the uploaded file." });
    }

    var results = new List<object>();
    var summaryRows = new List<object>();

    foreach (var row in rows)
    {
        var values = MonteCarloCalculator.RunSimulation(row, iterations);
        var summary = MonteCarloCalculator.GetSummary(values);

        var simulationRows = values.Select((value, index) => new
        {
            row = row.Name,
            iteration = index + 1,
            value
        }).ToList();

        results.AddRange(simulationRows);

        summaryRows.Add(new
        {
            name = row.Name,
            p20 = summary.P20,
            p50 = summary.P50,
            p80 = summary.P80
        });
    }

    return Results.Ok(new
    {
        iterations,
        summary = summaryRows,
        values = results
    });
});

app.MapFallbackToFile("index.html");

app.Run();

using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using StatementFile.Infrastructure.Configuration;

var builder = WebApplication.CreateBuilder(args);

// ── Services ──────────────────────────────────────────────────────────────────

builder.Services.AddControllers()
    .AddJsonOptions(opts =>
    {
        // Serialize enums as their integer values for cross-language compatibility.
        opts.JsonSerializerOptions.Converters.Clear();
    });

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title   = "StatementFile API",
        Version = "v1",
        Description =
            "REST API for configurable bank statement generation. " +
            "Replaces the 500-case switch statement in frmStatementFileExtn.runStatement() " +
            "with database-driven configuration stored in STAT_BANK_PRODUCT_CONFIG.",
    });
});

// Register all Infrastructure dependencies (Oracle, Email, Formatters, Handlers).
builder.Services.AddInfrastructure(builder.Configuration);

// CORS — allow the Blazor / React front-end to call this API.
var allowedOrigins = builder.Configuration["Cors:AllowedOrigins"]?.Split(',')
                    ?? new[] { "http://localhost:5000", "https://localhost:5001" };

builder.Services.AddCors(options => options.AddDefaultPolicy(policy =>
    policy.WithOrigins(allowedOrigins)
          .AllowAnyHeader()
          .AllowAnyMethod()));

// ── Pipeline ──────────────────────────────────────────────────────────────────

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "StatementFile API v1");
        c.RoutePrefix = string.Empty;   // Swagger UI at root
    });
}

app.UseCors();
app.MapControllers();

app.Run();

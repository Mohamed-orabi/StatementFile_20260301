using StatementFile.Infrastructure.Configuration;

var builder = WebApplication.CreateBuilder(args);

// ── Controllers + OpenAPI ──────────────────────────────────────────────────────
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(c =>
{
    c.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title   = "StatementFile API",
        Version = "v1",
        Description =
            "REST API for the StatementFile statement-generation platform.\n" +
            "Wraps the Domain → Application → Infrastructure layers behind HTTP endpoints\n" +
            "so the Blazor Presentation layer is fully decoupled from Oracle / Crystal Reports."
    });
});

// ── Composition root (singleton – one OracleConnectionFactory for the host) ───
builder.Services.AddSingleton(_ => DependencyInjection.Compose());

// ── CORS – allow the Blazor front-end ─────────────────────────────────────────
var allowedOrigins = builder.Configuration
    .GetSection("Cors:AllowedOrigins")
    .Get<string[]>() ?? new[] { "http://localhost:5000", "https://localhost:5001" };

builder.Services.AddCors(options =>
    options.AddDefaultPolicy(policy =>
        policy.WithOrigins(allowedOrigins)
              .AllowAnyHeader()
              .AllowAnyMethod()));

var app = builder.Build();

// ── Middleware ─────────────────────────────────────────────────────────────────
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c => c.SwaggerEndpoint("/swagger/v1/swagger.json", "StatementFile API v1"));
}

app.UseCors();
app.MapControllers();

app.Run();

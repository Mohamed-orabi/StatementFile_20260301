using StatementFile.Infrastructure.Configuration;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// Composition root is a singleton — one Oracle connection factory for the lifetime
// of the web host, exactly as in the original WinForms Program.Main().
builder.Services.AddSingleton(_ => DependencyInjection.Compose());

// Optional: enable OpenAPI/Swagger in development
builder.Services.AddEndpointsApiExplorer();

var app = builder.Build();

app.MapControllers();

app.Run();

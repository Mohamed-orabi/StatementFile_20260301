using System;
using StatementFile.Infrastructure.Configuration;
using StatementFile.Presentation.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

// MVC controllers serve the embedded REST API (e.g. BankConfigController at /api/bank-config).
builder.Services.AddControllers();

// Composition root is a singleton — one Oracle connection factory for the lifetime
// of the web host, exactly as in the original WinForms Program.Main().
builder.Services.AddSingleton(_ => DependencyInjection.Compose());

// AppState is scoped so each browser session gets its own login state.
builder.Services.AddScoped<AppState>();

// Typed HttpClient for the bank-config REST API consumed by Blazor pages.
// Base address must match the URL this application actually listens on
// (configured in appsettings.json → "ApiBaseUrl").
builder.Services.AddHttpClient<IBankProductConfigApiClient, BankProductConfigApiClient>(client =>
{
    var baseUrl = builder.Configuration["ApiBaseUrl"]
                  ?? throw new InvalidOperationException(
                      "ApiBaseUrl is not configured. " +
                      "Add it to appsettings.json, e.g. \"ApiBaseUrl\": \"http://localhost:5000\".");
    client.BaseAddress = new Uri(baseUrl);
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseStaticFiles();
app.UseAntiforgery();

// REST API endpoints (controllers)
app.MapControllers();

app.MapRazorComponents<StatementFile.Presentation.App>()
   .AddInteractiveServerRenderMode();

app.Run();

using System;
using StatementFile.Presentation.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

// AppState is scoped so each browser session gets its own login state.
builder.Services.AddScoped<AppState>();

// All Blazor pages communicate with the backend through StatementFile.Api.
// Base address must match the URL the Api project listens on
// (configured in appsettings.json → "ApiBaseUrl").
Uri apiBaseAddress = new Uri(
    builder.Configuration["ApiBaseUrl"]
    ?? throw new InvalidOperationException(
        "ApiBaseUrl is not configured. " +
        "Add it to appsettings.json, e.g. \"ApiBaseUrl\": \"http://localhost:5000\"."));

builder.Services.AddHttpClient<IAuthApiClient, AuthApiClient>(
    client => client.BaseAddress = apiBaseAddress);

builder.Services.AddHttpClient<IBankProductConfigApiClient, BankProductConfigApiClient>(
    client => client.BaseAddress = apiBaseAddress);

builder.Services.AddHttpClient<IStatementGenerationApiClient, StatementGenerationApiClient>(
    client => client.BaseAddress = apiBaseAddress);

builder.Services.AddHttpClient<IMerchantStatementApiClient, MerchantStatementApiClient>(
    client => client.BaseAddress = apiBaseAddress);

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<StatementFile.Presentation.App>()
   .AddInteractiveServerRenderMode();

app.Run();

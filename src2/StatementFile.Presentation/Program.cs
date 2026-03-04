using StatementFile.Presentation.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

// AppState is scoped so each browser session has its own login state.
builder.Services.AddScoped<AppState>();

// StatementApiClient is a typed HTTP client that calls StatementFile.Api.
// The base address is read from appsettings.json "StatementApiBaseUrl".
builder.Services.AddHttpClient<StatementApiClient>(client =>
{
    var baseUrl = builder.Configuration["StatementApiBaseUrl"]
                  ?? "http://localhost:5100";
    client.BaseAddress = new Uri(baseUrl);
    client.Timeout = TimeSpan.FromMinutes(30); // generation can take several minutes
});

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

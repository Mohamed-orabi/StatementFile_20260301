using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using StatementFile.Blazor.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// Typed HttpClient that calls the StatementFile API
builder.Services.AddHttpClient<StatementApiClient>(client =>
{
    var baseUrl = builder.Configuration["StatementApiBaseUrl"]
                  ?? "http://localhost:5100";
    client.BaseAddress = new Uri(baseUrl);
    client.Timeout     = TimeSpan.FromMinutes(30);
});

var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseAntiforgery();

app.MapRazorComponents<StatementFile.Blazor.Components.App>()
   .AddInteractiveServerRenderMode();

app.Run();

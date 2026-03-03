using StatementFile.Infrastructure.Configuration;
using StatementFile.Presentation.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents()
                .AddInteractiveServerComponents();

// Composition root is a singleton — one Oracle connection factory for the lifetime
// of the web host, exactly as in the original WinForms Program.Main().
builder.Services.AddSingleton(_ => DependencyInjection.Compose());

// AppState is scoped so each browser session gets its own login state.
builder.Services.AddScoped<AppState>();

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

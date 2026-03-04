using System;
using Microsoft.AspNetCore.Authentication.Cookies;
using StatementFile.Presentation.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddRazorComponents();

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath         = "/login";
        options.SlidingExpiration = true;
        options.ExpireTimeSpan    = TimeSpan.FromHours(8);
    });

builder.Services.AddAuthorization();
builder.Services.AddHttpContextAccessor();

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
app.UseAuthentication();
app.UseAuthorization();
app.UseAntiforgery();

// Logout endpoint — POST only so it can't be triggered by a GET link.
app.MapPost("/logout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/login");
});

app.MapRazorComponents<StatementFile.Presentation.App>();

app.Run();

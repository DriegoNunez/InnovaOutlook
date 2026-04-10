using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenLocalhost(3000, listenOptions =>
    {
        listenOptions.UseHttps();
    });
});

var app = builder.Build();

var repoRoot = Directory.GetParent(app.Environment.ContentRootPath)?.FullName
    ?? throw new InvalidOperationException("Unable to determine the repository root.");

var staticFiles = new PhysicalFileProvider(repoRoot);

app.UseDefaultFiles(new DefaultFilesOptions
{
    FileProvider = staticFiles
});

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = staticFiles,
    OnPrepareResponse = context =>
    {
        context.Context.Response.Headers.CacheControl = "no-store, no-cache, max-age=0";
        context.Context.Response.Headers.Pragma = "no-cache";
        context.Context.Response.Headers.Expires = "0";
    }
});

app.MapGet("/health", () => Results.Ok(new
{
    status = "ok",
    root = repoRoot
}));

app.Run();

using System.Globalization;
using Microsoft.AspNetCore.Localization;
using Newtonsoft.Json.Serialization;
using GuptaBrothers.Web.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromHours(4);
    options.Cookie.HttpOnly = true;
    options.Cookie.Path = "/";
    options.Cookie.Name = "GuptaBrothers-SESSION";
});

builder.Services.AddControllersWithViews()
    .AddNewtonsoftJson(opt => opt.SerializerSettings.ContractResolver = new DefaultContractResolver())
    .AddSessionStateTempDataProvider();

if (builder.Environment.IsDevelopment())
{
    builder.Services.AddRazorPages().AddRazorRuntimeCompilation();
}

// Register Gupta Brothers Application Services
builder.Services.AddSingleton<IQrCodeBarcodeService, QrCodeBarcodeService>();
builder.Services.AddSingleton<IExcelService, ExcelService>();
builder.Services.AddSingleton<IShippingLabelService, ShippingLabelService>();
builder.Services.AddSingleton<IPdfService, PdfService>();

var app = builder.Build();

var supportedCultures = new[] { new CultureInfo("en-IN"), new CultureInfo("en-US") };
app.UseRequestLocalization(new RequestLocalizationOptions
{
    DefaultRequestCulture = new RequestCulture("en-IN"),
    SupportedCultures = supportedCultures,
});

if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
}
else
{
    app.UseExceptionHandler("/ShippingLabel/Index");
    app.UseHsts();
}

app.UseStaticFiles();
app.UseRouting();
app.UseSession();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=ShippingLabel}/{action=Index}/{id?}");

app.Run();

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

if (args.Contains("--test-pdf"))
{
    var labelSvc = app.Services.GetRequiredService<IShippingLabelService>();
    var pdfSvc = app.Services.GetRequiredService<IPdfService>();
    var list = labelSvc.GetAll().Take(50).ToList();
    Console.WriteLine($"[PDF TEST] Generating PDF for {list.Count} labels...");
    var sw = System.Diagnostics.Stopwatch.StartNew();
    var bytes = pdfSvc.GenerateBatchLabelPdf(list);
    sw.Stop();
    Console.WriteLine($"[PDF TEST] Generated {bytes.Length} bytes in {sw.ElapsedMilliseconds} ms!");
    File.WriteAllBytes(Path.Combine(app.Environment.ContentRootPath, "App_Data", "test_50_labels.pdf"), bytes);
    return;
}

if (args.Contains("--import-final"))
{
    var excelSvc = app.Services.GetRequiredService<IExcelService>();
    var labelSvc = app.Services.GetRequiredService<IShippingLabelService>();
    var filePath = Path.Combine(app.Environment.ContentRootPath, "App_Data", "Uploaded_Final_Template.xlsx");
    if (File.Exists(filePath))
    {
        using var stream = File.OpenRead(filePath);
        var imported = excelSvc.ImportFromExcel(stream, labelSvc.GetSettings());
        Console.WriteLine($"[IMPORT TEST] Successfully imported: {imported.Count} records!");
        if (imported.Count > 0)
        {
            labelSvc.ClearAll();
            labelSvc.AddRange(imported);
            Console.WriteLine($"[IMPORT TEST] Saved {imported.Count} records to database/cache.");
            Console.WriteLine($"Sample 1: AWB={imported[0].AWBNumber}, Customer={imported[0].CustomerName}, Seller={imported[0].SoldBy}, City={imported[0].CustomerCity}, State={imported[0].CustomerState}, Pin={imported[0].CustomerPincode}");
            Console.WriteLine($"Sample 2: AWB={imported[1].AWBNumber}, Customer={imported[1].CustomerName}, Seller={imported[1].SoldBy}, City={imported[1].CustomerCity}, State={imported[1].CustomerState}, Pin={imported[1].CustomerPincode}");
            Console.WriteLine($"Last record: AWB={imported[^1].AWBNumber}, Customer={imported[^1].CustomerName}, Seller={imported[^1].SoldBy}");
        }
    }
    else
    {
        Console.WriteLine($"[IMPORT TEST] File not found: {filePath}");
    }
    return;
}

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

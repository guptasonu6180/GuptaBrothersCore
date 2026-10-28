# Gupta Brothers - Enterprise Shipping Label & Tax Invoice Suite

A modern, standalone .NET 10 web application designed for **Gupta Brothers** to import orders from Excel, generate dynamic vector **Code 128 Barcodes** and **High-Resolution QR Codes**, preview tax invoices, export data, and print thermal 4" × 6" / A6 shipping labels.

---

## 📂 Project Structure

```
D:\GuptaBrothers\
│
├── GuptaBrothers.sln             # Visual Studio Solution File
├── GuptaBrothers.slnx            # Modern .NET Solution File
├── README.md                     # Documentation & Setup Guide
│
└── src\
    └── GuptaBrothers.Web\        # Core Web Application (.NET 10)
        ├── Controllers\
        │   └── ShippingLabelController.cs
        ├── Models\
        │   └── ShippingLabelModel.cs
        ├── Services\
        │   ├── IQrCodeBarcodeService.cs
        │   ├── QrCodeBarcodeService.cs   # Base64 PNG + SVG QR & Code 128
        │   ├── IExcelService.cs
        │   ├── ExcelService.cs           # ClosedXML Import & Export
        │   ├── IShippingLabelService.cs
        │   └── ShippingLabelService.cs   # JSON Persistence & Defaults
        ├── Views\
        │   ├── Shared\_Layout.cshtml    # Futuristic Cyber Theme
        │   └── ShippingLabel\
        │       ├── Index.cshtml          # Dashboard, Live Filter & Actions
        │       ├── Print.cshtml          # Exact 4x6 / A6 Thermal Printout
        │       ├── ViewPdf.cshtml        # Full Label Preview
        │       ├── Create.cshtml         # Manual Label Entry
        │       ├── Edit.cshtml           # Edit Order
        │       └── Settings.cshtml       # Seller & QR Configuration
        └── wwwroot\                      # Bootstrap 5, Icons & Fonts
```

---

## ⚡ How to Run

### Method 1: Using .NET CLI (Command Prompt or PowerShell)

```powershell
cd D:\GuptaBrothers
dotnet run --project src\GuptaBrothers.Web
```

Open your browser at:
- **HTTP**: [http://localhost:5100](http://localhost:5100)
- **HTTPS**: [https://localhost:5101](https://localhost:5101)

### Method 2: Visual Studio
Double-click `D:\GuptaBrothers\GuptaBrothers.sln` to open directly in Visual Studio, then press `F5` or `Ctrl + F5`.

---

## 🌟 Key Capabilities

1. **Excel Import & Dynamic Mapping**:
   - Upload any `.xlsx` or `.xls` file.
   - Automatically maps columns (`AWB`, `Order No`, `Customer Name`, `Address`, `SKU`, `Gross Amount`, `GST`, etc.).
   - Automatically computes 3.0% IGST/GST, taxable value, and totals if not explicitly provided.

2. **Pixel-Perfect Thermal Print Layout**:
   - Matches the e-commerce standard format (Valmo / Meesho / Amazon / Flipkart style).
   - Formatted with `@media print` rules for thermal printers (100mm × 150mm) or standard A4/A6.

3. **Dynamic QR Code & Barcode Engine**:
   - **QR Code**: Rendered in high-definition Base64 PNG and vector SVG. Clicking any QR code opens the **Inspector Modal** with full details and a 1-click **Download QR Image** button.
   - **Barcode**: Pure C# vector Code 128 barcode generator with zero native or external dependencies.

4. **Sample Template Included**:
   - Click **"Sample Template"** on the dashboard to download `GuptaBrothers_Shipping_Template.xlsx` pre-loaded with the 7 sample orders ready to test.

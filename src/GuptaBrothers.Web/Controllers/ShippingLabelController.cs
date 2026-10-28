using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using GuptaBrothers.Web.Models;
using GuptaBrothers.Web.Services;

namespace GuptaBrothers.Web.Controllers
{
    public class ShippingLabelController : Controller
    {
        private readonly IShippingLabelService _labelService;
        private readonly IExcelService _excelService;
        private readonly IQrCodeBarcodeService _qrBarcodeService;
        private readonly IPdfService _pdfService;

        public ShippingLabelController(
            IShippingLabelService labelService,
            IExcelService excelService,
            IQrCodeBarcodeService qrBarcodeService,
            IPdfService pdfService)
        {
            _labelService = labelService;
            _excelService = excelService;
            _qrBarcodeService = qrBarcodeService;
            _pdfService = pdfService;
        }

        [HttpGet]
        public IActionResult Index(string? search, int page = 1, int pageSize = 50)
        {
            if (page < 1) page = 1;
            if (pageSize <= 0 && pageSize != -1) pageSize = 50;

            ViewBag.SearchTerm = search;
            var list = _labelService.GetAll(search);

            ViewBag.TotalCount = list.Count;
            ViewBag.TotalItems = list.Sum(x => x.Qty);
            ViewBag.TotalAmount = list.Sum(x => x.GrandTotal);
            ViewBag.TotalTaxes = list.Sum(x => x.TotalTaxes);

            int totalRecords = list.Count;
            bool showAll = (pageSize == -1 || pageSize >= totalRecords);
            int totalPages = showAll ? 1 : (int)Math.Ceiling((double)totalRecords / pageSize);
            if (totalPages < 1) totalPages = 1;
            if (page > totalPages) page = totalPages;

            var pagedList = showAll
                ? list
                : list.Skip((page - 1) * pageSize).Take(pageSize).ToList();

            // Populate QR & Barcode dynamically for the visible records
            Parallel.ForEach(pagedList, item => _labelService.EnsureQrAndBarcode(item));

            ViewBag.CurrentPage = page;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalPages = totalPages;

            return View(pagedList);
        }

        [HttpGet]
        public IActionResult DownloadTemplate()
        {
            var bytes = _excelService.GenerateSampleTemplate();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "GuptaBrothers_Shipping_Template.xlsx");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ImportExcel(IFormFile? excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["ErrorMessage"] = "Please select a valid Excel file (.xlsx) to upload.";
                return RedirectToAction(nameof(Index));
            }

            var extension = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls" && extension != ".csv")
            {
                TempData["ErrorMessage"] = "Unsupported file type. Please upload a .xlsx Excel workbook or .csv file.";
                return RedirectToAction(nameof(Index));
            }

            try
            {
                var settings = _labelService.GetSettings();
                using var stream = excelFile.OpenReadStream();
                var imported = _excelService.ImportFromFile(stream, excelFile.FileName, settings);

                if (imported.Count == 0)
                {
                    TempData["ErrorMessage"] = "No shipping label records could be read from the file. Please check column headers or download the sample template.";
                }
                else
                {
                    _labelService.AddRange(imported);
                    TempData["SuccessMessage"] = $"Successfully imported {imported.Count} shipping orders from '{excelFile.FileName}'.";
                }
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = $"Failed to parse Excel file: {ex.Message}";
            }

            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult ExportExcel()
        {
            var list = _labelService.GetAll();
            var bytes = _excelService.ExportToExcel(list);
            var filename = $"GuptaBrothers_Labels_{DateTime.Now:yyyyMMdd_HHmm}.xlsx";
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
        }

        [HttpGet, HttpPost]
        public IActionResult ExportPdf(string? ids)
        {
            List<ShippingLabelModel> labels;
            if (string.IsNullOrWhiteSpace(ids))
            {
                labels = _labelService.GetAll();
                Parallel.ForEach(labels, item => _labelService.EnsureQrAndBarcode(item));
            }
            else
            {
                var idArray = ids.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                labels = _labelService.GetByIds(idArray);
            }

            if (labels.Count == 0)
            {
                TempData["ErrorMessage"] = "No shipping labels found to export to PDF.";
                return RedirectToAction(nameof(Index));
            }

            var pdfBytes = _pdfService.GenerateBatchLabelPdf(labels);
            var filename = labels.Count == 1
                ? $"GuptaBrothers_Label_{labels[0].AWBNumber}.pdf"
                : $"GuptaBrothers_Labels_Batch_{DateTime.Now:yyyyMMdd_HHmm}.pdf";

            return File(pdfBytes, "application/pdf", filename);
        }

        [HttpGet]
        public IActionResult DownloadPdf(string id)
        {
            var label = _labelService.GetById(id);
            if (label == null)
            {
                TempData["ErrorMessage"] = $"Label '{id}' not found.";
                return RedirectToAction(nameof(Index));
            }

            var pdfBytes = _pdfService.GenerateSingleLabelPdf(label);
            var filename = $"GuptaBrothers_Label_{label.AWBNumber}.pdf";
            return File(pdfBytes, "application/pdf", filename);
        }

        [HttpGet, HttpPost]
        public IActionResult Print(string? ids)
        {
            List<ShippingLabelModel> labels;
            if (string.IsNullOrWhiteSpace(ids))
            {
                labels = _labelService.GetAll();
                Parallel.ForEach(labels, item => _labelService.EnsureQrAndBarcode(item));
            }
            else
            {
                var idArray = ids.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                labels = _labelService.GetByIds(idArray);
            }

            if (labels.Count == 0)
            {
                TempData["ErrorMessage"] = "No labels selected to print.";
                return RedirectToAction(nameof(Index));
            }

            return View(labels);
        }

        [HttpGet]
        public IActionResult ViewPdf(string id)
        {
            var label = _labelService.GetById(id);
            if (label == null)
            {
                TempData["ErrorMessage"] = $"Label '{id}' not found.";
                return RedirectToAction(nameof(Index));
            }

            return View(label);
        }

        [HttpGet]
        public IActionResult Create()
        {
            var settings = _labelService.GetSettings();
            var model = new ShippingLabelModel
            {
                AWBNumber = "VL" + DateTime.Now.ToString("yyMMdd") + new Random().Next(10000, 99999),
                SoldBy = settings.DefaultSoldBy,
                SellerAddress = settings.DefaultSellerAddress,
                SellerGSTIN = settings.DefaultSellerGSTIN,
                ReturnCompanyName = settings.DefaultReturnCompanyName,
                ReturnAddressLine1 = settings.DefaultReturnAddressLine1,
                ReturnAddressLine2 = settings.DefaultReturnAddressLine2,
                ReturnArea = settings.DefaultReturnArea,
                ReturnCityStatePin = settings.DefaultReturnCityStatePin,
                CourierName = settings.DefaultCourierName,
                HubCode = settings.DefaultHubCode
            };
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Create(ShippingLabelModel model)
        {
            if (string.IsNullOrWhiteSpace(model.CustomerName))
            {
                ModelState.AddModelError("CustomerName", "Customer name is required.");
            }

            if (!ModelState.IsValid)
            {
                return View(model);
            }

            _labelService.Add(model);
            TempData["SuccessMessage"] = $"Label for '{model.CustomerName}' (AWB: {model.AWBNumber}) created successfully.";
            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult Edit(string id)
        {
            var item = _labelService.GetById(id);
            if (item == null)
            {
                TempData["ErrorMessage"] = "Record not found.";
                return RedirectToAction(nameof(Index));
            }
            return View(item);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Edit(ShippingLabelModel model)
        {
            if (!ModelState.IsValid)
            {
                return View(model);
            }

            _labelService.Update(model);
            TempData["SuccessMessage"] = $"Shipping label '{model.AWBNumber}' updated successfully.";
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Delete(string id)
        {
            var deleted = _labelService.Delete(id);
            if (deleted)
            {
                TempData["SuccessMessage"] = "Shipping label deleted.";
            }
            else
            {
                TempData["ErrorMessage"] = "Could not find record to delete.";
            }
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ClearAll()
        {
            _labelService.ClearAll();
            TempData["SuccessMessage"] = "All shipping label records have been cleared.";
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ResetSample()
        {
            _labelService.ResetToDefaultSampleData();
            TempData["SuccessMessage"] = "Reset to Gupta Brothers sample demo data successfully.";
            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public IActionResult Settings()
        {
            var settings = _labelService.GetSettings();
            return View(settings);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult Settings(LabelSettingsModel model)
        {
            _labelService.SaveSettings(model);
            TempData["SuccessMessage"] = "Settings saved successfully.";
            return RedirectToAction(nameof(Settings));
        }

        [HttpGet]
        public IActionResult BarcodeSvg(string text)
        {
            var svg = _qrBarcodeService.GenerateBarcodeSvg(text);
            return Content(svg, "image/svg+xml");
        }

        [HttpGet]
        public IActionResult QrCodeSvg(string text)
        {
            var svg = _qrBarcodeService.GenerateQrCodeSvg(text);
            return Content(svg, "image/svg+xml");
        }
    }
}

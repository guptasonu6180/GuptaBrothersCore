using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GuptaBrothers.Web.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace GuptaBrothers.Web.Services
{
    public class PdfService : IPdfService
    {
        private readonly IQrCodeBarcodeService _qrBarcodeService;

        static PdfService()
        {
            QuestPDF.Settings.License = LicenseType.Community;
            QuestPDF.Settings.ThrowOnMissingFontFamilies = false;
        }

        public PdfService(IQrCodeBarcodeService qrBarcodeService)
        {
            _qrBarcodeService = qrBarcodeService;
        }

        public byte[] GenerateSingleLabelPdf(ShippingLabelModel label)
        {
            return GenerateBatchLabelPdf(new[] { label });
        }

        public byte[] GenerateBatchLabelPdf(IEnumerable<ShippingLabelModel> labels)
        {
            var labelList = labels.ToList();
            if (labelList.Count == 0)
            {
                return Array.Empty<byte>();
            }

            // Ensure QR and Barcode are populated for every label
            var qrBytesMap = new Dictionary<string, byte[]>();
            foreach (var item in labelList)
            {
                if (string.IsNullOrEmpty(item.QrCodeBase64))
                {
                    string qrData = !string.IsNullOrWhiteSpace(item.AWBNumber) ? item.AWBNumber.Trim() : (item.OrderNo ?? "").Trim();
                    item.QRCodeData = qrData;
                    item.QrCodeBase64 = _qrBarcodeService.GenerateQrCodePngBase64(qrData, pixelsPerModule: 10);
                }
                if (string.IsNullOrEmpty(item.BarcodeSvg))
                {
                    var barcodeText = !string.IsNullOrWhiteSpace(item.AWBNumber) ? item.AWBNumber.Trim() : (item.OrderNo ?? "").Trim();
                    item.BarcodeSvg = _qrBarcodeService.GenerateBarcodeSvg(barcodeText, height: 65, barWidth: 3);
                }

                if (!string.IsNullOrEmpty(item.QrCodeBase64))
                {
                    try
                    {
                        var base64 = item.QrCodeBase64;
                        var commaIdx = base64.IndexOf(',');
                        if (commaIdx >= 0) base64 = base64.Substring(commaIdx + 1);
                        qrBytesMap[item.Id] = Convert.FromBase64String(base64);
                    }
                    catch { }
                }
            }

            var document = Document.Create(container =>
            {
                foreach (var label in labelList)
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(18, Unit.Point);
                        page.PageColor(Colors.White);
                        page.DefaultTextStyle(x => x.FontFamily("Lato").FontSize(11).FontColor(Colors.Black));

                        page.Content().Column(col =>
                        {
                            // Half Label Container matching the exact 734x445 proportion
                            col.Item().Border(2).BorderColor(Colors.Black).Column(card =>
                            {
                                // Top Section: Customer / Return (Left 46%) & Logistics / Barcode (Right 54%)
                                card.Item().BorderBottom(2).BorderColor(Colors.Black).Row(topRow =>
                                {
                                    // Left Column
                                    topRow.RelativeItem(46).BorderRight(2).BorderColor(Colors.Black).Column(left =>
                                    {
                                        // Customer Address
                                        left.Item().BorderBottom(2).BorderColor(Colors.Black).Padding(6).Column(c =>
                                        {
                                            c.Item().Text("Customer Address").Bold().FontSize(12);
                                            c.Item().Text(label.CustomerName).Bold().FontSize(14.5f);
                                            if (!string.IsNullOrWhiteSpace(label.CustomerAddressLine1))
                                            {
                                                c.Item().Text(label.CustomerAddressLine1).FontSize(10.5f);
                                            }
                                            if (!string.IsNullOrWhiteSpace(label.CustomerAddressLine2))
                                            {
                                                c.Item().Text(label.CustomerAddressLine2).FontSize(10.5f);
                                            }
                                            var cParts = new List<string>();
                                            if (!string.IsNullOrWhiteSpace(label.CustomerCity)) cParts.Add(label.CustomerCity.Trim());
                                            if (!string.IsNullOrWhiteSpace(label.CustomerState) && !string.Equals(label.CustomerState.Trim(), label.CustomerCity?.Trim(), StringComparison.OrdinalIgnoreCase)) cParts.Add(label.CustomerState.Trim());
                                            if (!string.IsNullOrWhiteSpace(label.CustomerPincode)) cParts.Add(label.CustomerPincode.Trim());
                                            if (cParts.Count > 0)
                                            {
                                                c.Item().Text(string.Join(", ", cParts)).FontSize(10.5f);
                                            }
                                        });

                                        // Return Address
                                        left.Item().Padding(6).Column(r =>
                                        {
                                            r.Item().Text("If undelivered, return to:").Bold().FontSize(11);
                                            r.Item().Text((label.ReturnCompanyName ?? label.SoldBy ?? "").ToUpperInvariant()).Bold().FontSize(12.5f);
                                            if (!string.IsNullOrWhiteSpace(label.ReturnAddressLine1))
                                            {
                                                r.Item().Text(label.ReturnAddressLine1).FontSize(10f);
                                            }
                                            if (!string.IsNullOrWhiteSpace(label.ReturnAddressLine2))
                                            {
                                                r.Item().Text(label.ReturnAddressLine2).FontSize(10f);
                                            }
                                            if (!string.IsNullOrWhiteSpace(label.ReturnArea))
                                            {
                                                r.Item().Text(label.ReturnArea).FontSize(10f);
                                            }
                                            if (!string.IsNullOrWhiteSpace(label.ReturnCityStatePin))
                                            {
                                                r.Item().Text(label.ReturnCityStatePin).FontSize(10f);
                                            }
                                        });
                                    });

                                    // Right Column
                                    topRow.RelativeItem(54).Column(right =>
                                    {
                                        // Payment Banner
                                        right.Item().Background(Colors.Black).Padding(4).PaddingLeft(6)
                                            .Text(label.PaymentStatus).FontColor(Colors.White).Bold().FontSize(12f);

                                        right.Item().Padding(6).Column(mid =>
                                        {
                                            mid.Item().Row(mRow =>
                                            {
                                                mRow.RelativeItem().Column(log =>
                                                {
                                                    log.Item().Row(cr =>
                                                    {
                                                        cr.AutoItem().Text(label.CourierName).Bold().FontSize(18);
                                                        cr.AutoItem().PaddingLeft(5).Background(Colors.Black).Padding(2).PaddingHorizontal(4)
                                                            .Text(label.PickupBadgeText).FontColor(Colors.White).Bold().FontSize(10);
                                                        cr.AutoItem().PaddingLeft(4).Border(1.5f).BorderColor(Colors.Black).Padding(1.5f).PaddingHorizontal(5)
                                                            .Text(label.PickupDateBadge).Bold().FontSize(14);
                                                    });

                                                    log.Item().PaddingTop(4).Text(label.HubCode).ExtraBold().FontSize(21);

                                                    log.Item().PaddingTop(22).Column(rt =>
                                                    {
                                                        if (!string.IsNullOrWhiteSpace(label.RoutingLine1))
                                                            rt.Item().Text(label.RoutingLine1).Bold().FontSize(18);
                                                        if (!string.IsNullOrWhiteSpace(label.RoutingLine2))
                                                            rt.Item().Text(label.RoutingLine2).Bold().FontSize(18);
                                                        if (!string.IsNullOrWhiteSpace(label.RoutingLine3))
                                                            rt.Item().Text(label.RoutingLine3).Bold().FontSize(18);
                                                    });
                                                });

                                                // High-res QR code image (Strictly encodes only AWB Tracking Number)
                                                if (qrBytesMap.TryGetValue(label.Id, out var qrBytes))
                                                {
                                                    try
                                                    {
                                                        mRow.ConstantItem(112).Height(112).Image(qrBytes);
                                                    }
                                                    catch { }
                                                }
                                            });

                                            // AWB Number & Barcode
                                            mid.Item().PaddingTop(5).AlignCenter().Text(label.AWBNumber).Bold().FontSize(13.5f);
                                            if (!string.IsNullOrWhiteSpace(label.BarcodeSvg))
                                            {
                                                try
                                                {
                                                    mid.Item().AlignCenter().Width(240).Height(44).Svg(label.BarcodeSvg);
                                                }
                                                catch { }
                                            }
                                        });
                                    });
                                });

                                // Bottom Product Details Section
                                card.Item().Column(prod =>
                                {
                                    prod.Item().Padding(4).PaddingLeft(7).Text("Product Details").Bold().FontSize(12f);
                                    prod.Item().Table(table =>
                                    {
                                        table.ColumnsDefinition(cols =>
                                        {
                                            cols.RelativeColumn(35);
                                            cols.RelativeColumn(14);
                                            cols.RelativeColumn(9);
                                            cols.RelativeColumn(13);
                                            cols.RelativeColumn(29);
                                        });

                                        table.Header(h =>
                                        {
                                            h.Cell().Padding(1).PaddingLeft(7).Text("SKU").Bold().FontSize(11f);
                                            h.Cell().Padding(1).PaddingLeft(7).Text("Size").Bold().FontSize(11f);
                                            h.Cell().Padding(1).PaddingLeft(7).Text("Qty").Bold().FontSize(11f);
                                            h.Cell().Padding(1).PaddingLeft(7).Text("Color").Bold().FontSize(11f);
                                            h.Cell().Padding(1).PaddingLeft(7).Text("Order No.").Bold().FontSize(11f);
                                        });

                                        table.Cell().Padding(1).PaddingLeft(7).PaddingBottom(6).Text(label.SKU).FontSize(10.5f);
                                        table.Cell().Padding(1).PaddingLeft(7).PaddingBottom(6).Text(label.Size).FontSize(10.5f);
                                        table.Cell().Padding(1).PaddingLeft(7).PaddingBottom(6).Text(label.Qty.ToString()).FontSize(10.5f);
                                        table.Cell().Padding(1).PaddingLeft(7).PaddingBottom(6).Text(label.Color).FontSize(10.5f);
                                        table.Cell().Padding(1).PaddingLeft(7).PaddingBottom(6).Text(label.OrderNo).FontSize(10.5f);
                                    });
                                });
                            });
                        });
                    });
                }
            });

            using var ms = new MemoryStream();
            document.GeneratePdf(ms);
            return ms.ToArray();
        }
    }
}

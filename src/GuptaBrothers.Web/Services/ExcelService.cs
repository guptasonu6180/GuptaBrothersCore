using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using GuptaBrothers.Web.Models;

namespace GuptaBrothers.Web.Services
{
    public class ExcelService : IExcelService
    {
        private static readonly string[] TemplateHeaders = new string[]
        {
            "Date",
            "Customer Name",
            "Customer Address",
            "Return / Seller Name",
            "Return / Seller Address",
            "UFC Name",
            "Web QR Code Text",
            "Tracking Number",
            "SKU",
            "Size",
            "Qty",
            "Color",
            "Order No",
            "Payment Instruction"
        };

        public byte[] GenerateSampleTemplate()
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("ShippingLabels");

            // Row 1: Header Title
            var titleCell = worksheet.Cell(1, 1);
            titleCell.Value = "Shipping Label Data Export";
            worksheet.Range(1, 1, 1, TemplateHeaders.Length).Merge();
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            titleCell.Style.Font.FontColor = XLColor.White;
            titleCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f172a");
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(1).Height = 32;

            // Row 2: Subtitle
            var subCell = worksheet.Cell(2, 1);
            subCell.Value = "Extracted from Shipping Label | Courier: Valmo Pickup";
            worksheet.Range(2, 1, 2, TemplateHeaders.Length).Merge();
            subCell.Style.Font.FontSize = 10;
            subCell.Style.Font.FontColor = XLColor.FromHtml("#cbd5e1");
            subCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
            subCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(2).Height = 20;

            // Row 3: Spacer row
            worksheet.Row(3).Height = 8;

            // Row 4: Column Headers
            for (int col = 0; col < TemplateHeaders.Length; col++)
            {
                var cell = worksheet.Cell(4, col + 1);
                cell.Value = TemplateHeaders[col];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 10.5;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f172a");
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            worksheet.Row(4).Height = 26;

            // Row 5+: Sample Rows
            var sampleRows = GetSampleRecords();
            int rowIdx = 5;
            foreach (var item in sampleRows)
            {
                var r = worksheet.Row(rowIdx);
                r.Height = 48;

                // 1. Date
                r.Cell(1).SetValue(item.PickupDateBadge);
                r.Cell(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 2. Customer Name
                r.Cell(2).SetValue(item.CustomerName);
                r.Cell(2).Style.Font.Bold = true;

                // 3. Customer Address
                r.Cell(3).SetValue(FormatCustomerAddress(item));
                r.Cell(3).Style.Alignment.WrapText = true;

                // 4. Return / Seller Name
                r.Cell(4).SetValue(string.IsNullOrWhiteSpace(item.ReturnCompanyName) ? item.SoldBy : item.ReturnCompanyName);

                // 5. Return / Seller Address
                r.Cell(5).SetValue(FormatReturnAddress(item));
                r.Cell(5).Style.Alignment.WrapText = true;

                // 6. UFC Name
                r.Cell(6).SetValue(item.HubCode);
                r.Cell(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                r.Cell(6).Style.Font.Bold = true;

                // 7. Web QR Code Text (multiline)
                r.Cell(7).SetValue(FormatWebQrText(item));
                r.Cell(7).Style.Alignment.WrapText = true;
                r.Cell(7).Style.Font.Bold = true;

                // 8. Tracking Number
                r.Cell(8).SetValue(item.AWBNumber);
                r.Cell(8).Style.Font.Bold = true;
                r.Cell(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 9. SKU
                r.Cell(9).SetValue(item.SKU);
                r.Cell(9).Style.Font.Bold = true;

                // 10. Size
                r.Cell(10).SetValue(item.Size);
                r.Cell(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 11. Qty
                r.Cell(11).SetValue(item.Qty);
                r.Cell(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 12. Color
                r.Cell(12).SetValue(item.Color);
                r.Cell(12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 13. Order No
                r.Cell(13).SetValue(item.OrderNo);

                // 14. Payment Instruction
                r.Cell(14).SetValue(item.PaymentStatus);

                for (int c = 1; c <= TemplateHeaders.Length; c++)
                {
                    r.Cell(c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    r.Cell(c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    r.Cell(c).Style.Border.OutsideBorderColor = XLColor.FromHtml("#cbd5e1");
                }

                rowIdx++;
            }

            // Fixed Proportional Column Widths matching sample template
            worksheet.Column(1).Width = 12;   // Date
            worksheet.Column(2).Width = 20;   // Customer Name
            worksheet.Column(3).Width = 40;   // Customer Address
            worksheet.Column(4).Width = 24;   // Return / Seller Name
            worksheet.Column(5).Width = 45;   // Return / Seller Address
            worksheet.Column(6).Width = 14;   // UFC Name
            worksheet.Column(7).Width = 18;   // Web QR Code Text
            worksheet.Column(8).Width = 22;   // Tracking Number
            worksheet.Column(9).Width = 32;   // SKU
            worksheet.Column(10).Width = 14;  // Size
            worksheet.Column(11).Width = 10;  // Qty
            worksheet.Column(12).Width = 14;  // Color
            worksheet.Column(13).Width = 26;  // Order No
            worksheet.Column(14).Width = 36;  // Payment Instruction

            using var memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);
            return memoryStream.ToArray();
        }

        public byte[] ExportToExcel(IEnumerable<ShippingLabelModel> labels)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Labels");

            // Row 1: Header Title
            var titleCell = worksheet.Cell(1, 1);
            titleCell.Value = "Shipping Label Data Export";
            worksheet.Range(1, 1, 1, TemplateHeaders.Length).Merge();
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.FontSize = 14;
            titleCell.Style.Font.FontColor = XLColor.White;
            titleCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f172a");
            titleCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titleCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(1).Height = 32;

            // Row 2: Subtitle
            var subCell = worksheet.Cell(2, 1);
            subCell.Value = "Extracted from Shipping Label | Courier: Valmo Pickup";
            worksheet.Range(2, 1, 2, TemplateHeaders.Length).Merge();
            subCell.Style.Font.FontSize = 10;
            subCell.Style.Font.FontColor = XLColor.FromHtml("#cbd5e1");
            subCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1e293b");
            subCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            subCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            worksheet.Row(2).Height = 20;

            // Row 3: Spacer row
            worksheet.Row(3).Height = 8;

            // Row 4: Column Headers
            for (int col = 0; col < TemplateHeaders.Length; col++)
            {
                var cell = worksheet.Cell(4, col + 1);
                cell.Value = TemplateHeaders[col];
                cell.Style.Font.Bold = true;
                cell.Style.Font.FontSize = 10.5;
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#0f172a");
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            }
            worksheet.Row(4).Height = 26;

            int rowIdx = 5;
            foreach (var item in labels)
            {
                var r = worksheet.Row(rowIdx);
                r.Height = 48;

                // 1. Date
                r.Cell(1).SetValue(item.PickupDateBadge);
                r.Cell(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 2. Customer Name
                r.Cell(2).SetValue(item.CustomerName);
                r.Cell(2).Style.Font.Bold = true;

                // 3. Customer Address
                r.Cell(3).SetValue(FormatCustomerAddress(item));
                r.Cell(3).Style.Alignment.WrapText = true;

                // 4. Return / Seller Name
                r.Cell(4).SetValue(string.IsNullOrWhiteSpace(item.ReturnCompanyName) ? item.SoldBy : item.ReturnCompanyName);

                // 5. Return / Seller Address
                r.Cell(5).SetValue(FormatReturnAddress(item));
                r.Cell(5).Style.Alignment.WrapText = true;

                // 6. UFC Name
                r.Cell(6).SetValue(item.HubCode);
                r.Cell(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                r.Cell(6).Style.Font.Bold = true;

                // 7. Web QR Code Text
                r.Cell(7).SetValue(FormatWebQrText(item));
                r.Cell(7).Style.Alignment.WrapText = true;
                r.Cell(7).Style.Font.Bold = true;

                // 8. Tracking Number
                r.Cell(8).SetValue(item.AWBNumber);
                r.Cell(8).Style.Font.Bold = true;
                r.Cell(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 9. SKU
                r.Cell(9).SetValue(item.SKU);
                r.Cell(9).Style.Font.Bold = true;

                // 10. Size
                r.Cell(10).SetValue(item.Size);
                r.Cell(10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 11. Qty
                r.Cell(11).SetValue(item.Qty);
                r.Cell(11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 12. Color
                r.Cell(12).SetValue(item.Color);
                r.Cell(12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // 13. Order No
                r.Cell(13).SetValue(item.OrderNo);

                // 14. Payment Instruction
                r.Cell(14).SetValue(item.PaymentStatus);

                for (int c = 1; c <= TemplateHeaders.Length; c++)
                {
                    r.Cell(c).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    r.Cell(c).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    r.Cell(c).Style.Border.OutsideBorderColor = XLColor.FromHtml("#cbd5e1");
                }

                rowIdx++;
            }

            worksheet.Column(1).Width = 12;
            worksheet.Column(2).Width = 20;
            worksheet.Column(3).Width = 40;
            worksheet.Column(4).Width = 24;
            worksheet.Column(5).Width = 45;
            worksheet.Column(6).Width = 14;
            worksheet.Column(7).Width = 18;
            worksheet.Column(8).Width = 22;
            worksheet.Column(9).Width = 32;
            worksheet.Column(10).Width = 14;
            worksheet.Column(11).Width = 10;
            worksheet.Column(12).Width = 14;
            worksheet.Column(13).Width = 26;
            worksheet.Column(14).Width = 36;

            using var memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);
            return memoryStream.ToArray();
        }

        public List<ShippingLabelModel> ImportFromExcel(Stream stream, LabelSettingsModel? defaultSettings = null)
        {
            var results = new List<ShippingLabelModel>();
            var defaults = defaultSettings ?? new LabelSettingsModel();

            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            if (worksheet == null) return results;

            var lastRowUsed = worksheet.LastRowUsed();
            if (lastRowUsed == null) return results;
            int lastRowNum = lastRowUsed.RowNumber();

            // Search first 15 rows to find the header row containing key columns
            int headerRowIdx = 1;
            for (int r = 1; r <= Math.Min(15, lastRowNum); r++)
            {
                var row = worksheet.Row(r);
                int matchCount = 0;
                int cols = Math.Min(35, worksheet.ColumnsUsed().Count());
                for (int c = 1; c <= cols; c++)
                {
                    var text = row.Cell(c).GetString();
                    if (IsKnownHeader(text))
                    {
                        matchCount++;
                    }
                }

                if (matchCount >= 2)
                {
                    headerRowIdx = r;
                    break;
                }
            }

            var headerRow = worksheet.Row(headerRowIdx);
            var colIndexMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int col = 1; col <= worksheet.ColumnsUsed().Count(); col++)
            {
                var val = headerRow.Cell(col).GetString().Trim();
                if (!string.IsNullOrEmpty(val))
                {
                    var norm = NormalizeKey(val);
                    if (!string.IsNullOrEmpty(norm) && !colIndexMap.ContainsKey(norm))
                    {
                        colIndexMap[norm] = col;
                    }
                }
            }

            string GetVal(IXLRow row, params string[] possibleHeaders)
            {
                foreach (var h in possibleHeaders)
                {
                    var norm = NormalizeKey(h);
                    if (colIndexMap.TryGetValue(norm, out int colIndex))
                    {
                        var cell = row.Cell(colIndex);
                        if (cell == null || cell.IsEmpty()) continue;
                        try
                        {
                            var s = cell.GetFormattedString()?.Trim();
                            if (!string.IsNullOrEmpty(s)) return s;
                        }
                        catch { }
                        var fallback = cell.GetString()?.Trim();
                        if (!string.IsNullOrEmpty(fallback)) return fallback;
                    }
                }
                return "";
            }

            int GetInt(IXLRow row, params string[] possibleHeaders)
            {
                string s = GetVal(row, possibleHeaders);
                if (string.IsNullOrWhiteSpace(s)) return 1;
                if (int.TryParse(s, out int result)) return result;
                return 1;
            }

            for (int r = headerRowIdx + 1; r <= lastRowNum; r++)
            {
                var row = worksheet.Row(r);
                if (row.IsEmpty()) continue;

                var awb = GetVal(row, "waybillnumber", "waybill", "waybillno", "trackingnumber", "tracking no", "awb", "awbnumber", "tracking", "airwaybill");
                var customerName = GetVal(row, "cosutmername", "customername", "cosutmer", "customer", "buyername", "buyer", "recipientname", "recipient", "name", "clientname");
                var orderNo = GetVal(row, "orderid", "orderno", "order number", "order_no", "order", "purchaseorderno", "pono");

                if (string.IsNullOrWhiteSpace(awb) && string.IsNullOrWhiteSpace(customerName) && string.IsNullOrWhiteSpace(orderNo))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(awb))
                {
                    awb = "VL" + DateTime.Now.ToString("yyMMdd") + new Random().Next(10000, 99999);
                }

                var sellerName = GetVal(row, "sellername", "seller", "soldby", "returncompanyname", "returnsellername", "returnseller", "returnname", "vendorname", "vendor");
                var sellerAddr = GetVal(row, "selleraddress", "returnselleraddress", "returnaddress", "selleraddr", "returnaddr");
                var custAddr = GetVal(row, "cosutmeraddress", "cosutmer_address", "customeraddress", "deliveryaddress", "shippingaddress", "fulladdress", "address");
                var hubCode = GetVal(row, "hubcode", "hub", "ufcname", "ufc", "ufccode", "destinationhub");
                var pickupDate = GetVal(row, "date", "pickupdate", "pickup date", "pickupdatebadge", "orderdate", "batch");
                var payment = GetVal(row, "paymentinstruction", "paymentstatus", "payment", "paymenttype", "codprepaid", "cod");
                var sku = GetVal(row, "sku", "productcode", "itemcode", "itemdescription", "product");
                var size = GetVal(row, "size", "productsize");
                var qty = GetInt(row, "qty", "quantity", "itemqty");
                var color = GetVal(row, "color", "colour");

                var label = new ShippingLabelModel
                {
                    AWBNumber = awb,
                    QRCodeData = awb, // Strictly only tracking number in QR code
                    CourierName = defaults.DefaultCourierName,
                    PickupDateBadge = ValueOrDefault(pickupDate, "01/11"),
                    HubCode = ValueOrDefault(hubCode, defaults.DefaultHubCode),
                    PaymentStatus = ValueOrDefault(payment, "Check the payable amount on the app"),
                    CustomerName = customerName,
                    SKU = ValueOrDefault(sku, "SKU-" + (awb.Length >= 5 ? awb.Substring(awb.Length - 5) : new Random().Next(1000, 9999).ToString())),
                    Size = ValueOrDefault(size, "Free Size"),
                    Qty = qty > 0 ? qty : 1,
                    Color = ValueOrDefault(color, "NA"),
                    OrderNo = ValueOrDefault(orderNo, "ORD-" + new Random().Next(10000000, 99999999))
                };

                // Parse Web QR Code Text (multiline routing codes: line 1, line 2, line 3)
                var webQrText = GetVal(row, "webqrcodetext", "webqrcode", "routing", "routinglines");
                ParseRoutingLines(webQrText, label);

                // Parse Customer Address
                if (!string.IsNullOrWhiteSpace(custAddr))
                {
                    ParseCustomerAddress(custAddr, label);
                }
                else
                {
                    label.CustomerAddressLine1 = GetVal(row, "addressline1", "address 1", "street");
                    label.CustomerAddressLine2 = GetVal(row, "addressline2", "address 2", "locality", "landmark");
                    label.CustomerCity = GetVal(row, "city", "customercity", "district");
                    label.CustomerState = GetVal(row, "state", "customerstate", "province");
                    label.CustomerPincode = GetVal(row, "pincode", "customerpincode", "pin", "zip", "postalcode");
                }

                // Parse Return / Seller Name & Address
                label.ReturnCompanyName = ValueOrDefault(sellerName, defaults.DefaultReturnCompanyName);
                label.SoldBy = ValueOrDefault(sellerName, defaults.DefaultSoldBy);

                if (!string.IsNullOrWhiteSpace(sellerAddr))
                {
                    ParseReturnAddress(sellerAddr, label);
                }
                else
                {
                    label.ReturnAddressLine1 = ValueOrDefault(GetVal(row, "returnaddressline1"), defaults.DefaultReturnAddressLine1);
                    label.ReturnAddressLine2 = ValueOrDefault(GetVal(row, "returnaddressline2"), defaults.DefaultReturnAddressLine2);
                    label.ReturnArea = ValueOrDefault(GetVal(row, "returnarea"), defaults.DefaultReturnArea);
                    label.ReturnCityStatePin = ValueOrDefault(GetVal(row, "returncitystatepin"), defaults.DefaultReturnCityStatePin);
                    label.SellerAddress = ValueOrDefault(GetVal(row, "selleraddress"), defaults.DefaultSellerAddress);
                }

                label.SellerGSTIN = defaults.DefaultSellerGSTIN;
                label.AutoCalculate();
                results.Add(label);
            }

            return results;
        }

        public List<ShippingLabelModel> ImportFromFile(Stream stream, string fileName, LabelSettingsModel? defaultSettings = null)
        {
            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            if (ext == ".csv")
            {
                return ImportFromCsv(stream, defaultSettings);
            }
            return ImportFromExcel(stream, defaultSettings);
        }

        public List<ShippingLabelModel> ImportFromCsv(Stream stream, LabelSettingsModel? defaultSettings = null)
        {
            var results = new List<ShippingLabelModel>();
            var defaults = defaultSettings ?? new LabelSettingsModel();

            var records = ParseCsvRecords(stream);
            if (records.Count < 2) return results;

            // Find header record among first 10 records
            int headerIdx = 0;
            for (int i = 0; i < Math.Min(10, records.Count); i++)
            {
                int matchCount = records[i].Count(cell => IsKnownHeader(cell));
                if (matchCount >= 2)
                {
                    headerIdx = i;
                    break;
                }
            }

            var headerRow = records[headerIdx];
            var colIndexMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerRow.Count; i++)
            {
                var h = headerRow[i].Trim();
                if (!string.IsNullOrEmpty(h))
                {
                    var norm = NormalizeKey(h);
                    if (!string.IsNullOrEmpty(norm) && !colIndexMap.ContainsKey(norm))
                    {
                        colIndexMap[norm] = i;
                    }
                }
            }

            string GetVal(List<string> row, params string[] possibleHeaders)
            {
                foreach (var h in possibleHeaders)
                {
                    var norm = NormalizeKey(h);
                    if (colIndexMap.TryGetValue(norm, out int idx) && idx < row.Count)
                    {
                        var val = row[idx]?.Trim();
                        if (!string.IsNullOrEmpty(val)) return val;
                    }
                }
                return "";
            }

            int GetInt(List<string> row, params string[] possibleHeaders)
            {
                string s = GetVal(row, possibleHeaders);
                if (string.IsNullOrWhiteSpace(s)) return 1;
                if (int.TryParse(s, out int result)) return result;
                return 1;
            }

            for (int i = headerIdx + 1; i < records.Count; i++)
            {
                var row = records[i];
                if (row.Count == 0 || row.All(string.IsNullOrWhiteSpace)) continue;

                var awb = GetVal(row, "waybillnumber", "waybill", "waybillno", "trackingnumber", "tracking no", "awb", "awbnumber", "tracking", "airwaybill");
                var customerName = GetVal(row, "cosutmername", "customername", "cosutmer", "customer", "buyername", "buyer", "recipientname", "recipient", "name", "clientname");
                var orderNo = GetVal(row, "orderid", "orderno", "order number", "order_no", "order", "purchaseorderno", "pono");

                if (string.IsNullOrWhiteSpace(awb) && string.IsNullOrWhiteSpace(customerName) && string.IsNullOrWhiteSpace(orderNo))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(awb))
                {
                    awb = "VL" + DateTime.Now.ToString("yyMMdd") + new Random().Next(10000, 99999);
                }

                var sellerName = GetVal(row, "sellername", "seller", "soldby", "returncompanyname", "returnsellername", "returnseller", "returnname", "vendorname", "vendor");
                var sellerAddr = GetVal(row, "selleraddress", "returnselleraddress", "returnaddress", "selleraddr", "returnaddr");
                var custAddr = GetVal(row, "cosutmeraddress", "cosutmer_address", "customeraddress", "deliveryaddress", "shippingaddress", "fulladdress", "address");
                var hubCode = GetVal(row, "hubcode", "hub", "ufcname", "ufc", "ufccode", "destinationhub");
                var pickupDate = GetVal(row, "date", "pickupdate", "pickup date", "pickupdatebadge", "orderdate", "batch");
                var payment = GetVal(row, "paymentinstruction", "paymentstatus", "payment", "paymenttype", "codprepaid", "cod");
                var sku = GetVal(row, "sku", "productcode", "itemcode", "itemdescription", "product");
                var size = GetVal(row, "size", "productsize");
                var qty = GetInt(row, "qty", "quantity", "itemqty");
                var color = GetVal(row, "color", "colour");

                var label = new ShippingLabelModel
                {
                    AWBNumber = awb,
                    QRCodeData = awb, // Strictly only tracking number
                    CourierName = defaults.DefaultCourierName,
                    PickupDateBadge = ValueOrDefault(pickupDate, "01/11"),
                    HubCode = ValueOrDefault(hubCode, defaults.DefaultHubCode),
                    PaymentStatus = ValueOrDefault(payment, "Check the payable amount on the app"),
                    CustomerName = customerName,
                    SKU = ValueOrDefault(sku, "SKU-" + (awb.Length >= 5 ? awb.Substring(awb.Length - 5) : new Random().Next(1000, 9999).ToString())),
                    Size = ValueOrDefault(size, "Free Size"),
                    Qty = qty > 0 ? qty : 1,
                    Color = ValueOrDefault(color, "NA"),
                    OrderNo = ValueOrDefault(orderNo, "ORD-" + new Random().Next(10000000, 99999999))
                };

                // Parse Web QR Code Text
                var webQrText = GetVal(row, "webqrcodetext", "webqrcode", "routing", "routinglines");
                ParseRoutingLines(webQrText, label);

                // Parse Customer Address
                if (!string.IsNullOrWhiteSpace(custAddr))
                {
                    ParseCustomerAddress(custAddr, label);
                }
                else
                {
                    label.CustomerAddressLine1 = GetVal(row, "addressline1", "address 1", "street");
                    label.CustomerAddressLine2 = GetVal(row, "addressline2", "address 2", "locality", "landmark");
                    label.CustomerCity = GetVal(row, "city", "customercity", "district");
                    label.CustomerState = GetVal(row, "state", "customerstate", "province");
                    label.CustomerPincode = GetVal(row, "pincode", "customerpincode", "pin", "zip", "postalcode");
                }

                // Parse Return / Seller Name & Address
                label.ReturnCompanyName = ValueOrDefault(sellerName, defaults.DefaultReturnCompanyName);
                label.SoldBy = ValueOrDefault(sellerName, defaults.DefaultSoldBy);

                if (!string.IsNullOrWhiteSpace(sellerAddr))
                {
                    ParseReturnAddress(sellerAddr, label);
                }
                else
                {
                    label.ReturnAddressLine1 = ValueOrDefault(GetVal(row, "returnaddressline1"), defaults.DefaultReturnAddressLine1);
                    label.ReturnAddressLine2 = ValueOrDefault(GetVal(row, "returnaddressline2"), defaults.DefaultReturnAddressLine2);
                    label.ReturnArea = ValueOrDefault(GetVal(row, "returnarea"), defaults.DefaultReturnArea);
                    label.ReturnCityStatePin = ValueOrDefault(GetVal(row, "returncitystatepin"), defaults.DefaultReturnCityStatePin);
                    label.SellerAddress = ValueOrDefault(GetVal(row, "selleraddress"), defaults.DefaultSellerAddress);
                }

                label.SellerGSTIN = defaults.DefaultSellerGSTIN;
                label.AutoCalculate();
                results.Add(label);
            }

            return results;
        }

        public static string NormalizeKey(string? header)
        {
            if (string.IsNullOrWhiteSpace(header)) return "";
            var sb = new System.Text.StringBuilder();
            foreach (char c in header)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

        private static bool IsKnownHeader(string? header)
        {
            if (string.IsNullOrWhiteSpace(header)) return false;
            var norm = NormalizeKey(header);
            return norm.Contains("waybill") ||
                   norm.Contains("tracking") ||
                   norm.Contains("awb") ||
                   norm.Contains("order") ||
                   norm.Contains("seller") ||
                   norm.Contains("cosutmer") ||
                   norm.Contains("customer") ||
                   norm.Contains("hub") ||
                   norm.Contains("ufc") ||
                   norm.Contains("sku") ||
                   norm.Contains("product") ||
                   norm.Contains("item") ||
                   norm.Contains("address") ||
                   norm.Contains("date") ||
                   norm.Contains("qty") ||
                   norm.Contains("quantity") ||
                   norm.Contains("payment") ||
                   norm.Contains("return") ||
                   norm.Contains("soldby");
        }

        private static void ParseRoutingLines(string webQrText, ShippingLabelModel label)
        {
            if (string.IsNullOrWhiteSpace(webQrText))
            {
                label.RoutingLine1 = "W1/S11S";
                label.RoutingLine2 = "E1/BGUS";
                label.RoutingLine3 = "D-1/TQS";
                return;
            }

            var lines = webQrText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrEmpty(l))
                .ToArray();

            if (lines.Length >= 1) label.RoutingLine1 = lines[0];
            if (lines.Length >= 2) label.RoutingLine2 = lines[1];
            if (lines.Length >= 3) label.RoutingLine3 = lines[2];
        }

        private static void ParseCustomerAddress(string fullAddress, ShippingLabelModel label)
        {
            if (string.IsNullOrWhiteSpace(fullAddress)) return;

            fullAddress = fullAddress.Trim().TrimEnd(',');

            // Extract 6-digit pincode if present
            var pinMatch = Regex.Match(fullAddress, @"\b([1-9][0-9]{5})\b");
            if (pinMatch.Success)
            {
                label.CustomerPincode = pinMatch.Value;
            }

            var parts = fullAddress.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrEmpty(p))
                .ToList();

            if (parts.Count == 0)
            {
                label.CustomerAddressLine1 = fullAddress;
                return;
            }

            // Remove standalone pincode from last part if matched
            if (parts.Count > 1 && Regex.IsMatch(parts[parts.Count - 1], @"^\d{6}$"))
            {
                parts.RemoveAt(parts.Count - 1);
            }

            if (parts.Count == 1)
            {
                label.CustomerAddressLine1 = parts[0];
            }
            else if (parts.Count == 2)
            {
                label.CustomerAddressLine1 = parts[0];
                label.CustomerAddressLine2 = parts[1];
                var words = parts[1].Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 1)
                {
                    label.CustomerCity = words[0];
                }
                else if (words.Length >= 2)
                {
                    label.CustomerCity = words[words.Length - 2];
                    label.CustomerState = words[words.Length - 1];
                }
            }
            else if (parts.Count == 3)
            {
                label.CustomerAddressLine1 = parts[0];
                label.CustomerAddressLine2 = parts[1];
                label.CustomerCity = parts[2];
            }
            else
            {
                label.CustomerAddressLine1 = parts[0];
                label.CustomerAddressLine2 = string.Join(", ", parts.Skip(1).Take(parts.Count - 3));
                if (string.IsNullOrWhiteSpace(label.CustomerAddressLine2))
                {
                    label.CustomerAddressLine2 = parts[1];
                }
                label.CustomerCity = parts[parts.Count - 2];
                label.CustomerState = parts[parts.Count - 1];
            }
        }

        private static void ParseReturnAddress(string fullAddress, ShippingLabelModel label)
        {
            if (string.IsNullOrWhiteSpace(fullAddress)) return;
            label.SellerAddress = fullAddress.Trim().TrimEnd(',');

            var pinMatch = Regex.Match(fullAddress, @"\b([1-9][0-9]{5})\b");
            string pin = pinMatch.Success ? pinMatch.Value : "395006";

            var parts = fullAddress.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(p => p.Trim())
                .Where(p => !string.IsNullOrEmpty(p))
                .ToList();

            if (parts.Count == 0) return;

            var streetParts = parts.Where(p =>
                !Regex.IsMatch(p, @"^(pincode\s*)?\d{6}$", RegexOptions.IgnoreCase) &&
                !string.Equals(p, "Gujarat", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(p, "Surat", StringComparison.OrdinalIgnoreCase)
            ).ToList();

            if (streetParts.Count >= 3)
            {
                label.ReturnAddressLine1 = string.Join(", ", streetParts.Take(streetParts.Count - 2));
                label.ReturnAddressLine2 = streetParts[streetParts.Count - 2];
                label.ReturnArea = streetParts[streetParts.Count - 1];
            }
            else if (streetParts.Count == 2)
            {
                label.ReturnAddressLine1 = streetParts[0];
                label.ReturnAddressLine2 = streetParts[1];
                label.ReturnArea = "Kapodra";
            }
            else if (streetParts.Count == 1)
            {
                label.ReturnAddressLine1 = streetParts[0];
                label.ReturnArea = "Kapodra";
            }
            else
            {
                label.ReturnAddressLine1 = fullAddress;
            }

            label.ReturnCityStatePin = $"Surat, Gujarat, {pin}";
        }

        private static string FormatCustomerAddress(ShippingLabelModel label)
        {
            var parts = new List<string>
            {
                label.CustomerAddressLine1,
                label.CustomerAddressLine2,
                label.CustomerCity,
                label.CustomerState,
                label.CustomerPincode
            }.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim().TrimEnd(','));

            return string.Join(", ", parts);
        }

        private static string FormatReturnAddress(ShippingLabelModel label)
        {
            var parts = new List<string>
            {
                label.ReturnAddressLine1,
                label.ReturnAddressLine2,
                label.ReturnArea,
                label.ReturnCityStatePin
            }.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim().TrimEnd(',')).ToList();

            if (parts.Count > 0)
            {
                return string.Join(", ", parts);
            }

            return label.SellerAddress ?? "";
        }

        private static string FormatWebQrText(ShippingLabelModel label)
        {
            var lines = new[] { label.RoutingLine1, label.RoutingLine2, label.RoutingLine3 }
                .Where(s => !string.IsNullOrWhiteSpace(s));
            return string.Join("\n", lines);
        }

        private static List<List<string>> ParseCsvRecords(Stream stream)
        {
            var records = new List<List<string>>();
            using var reader = new StreamReader(stream, System.Text.Encoding.UTF8, true, leaveOpen: true);

            var currentRecord = new List<string>();
            var currentField = new System.Text.StringBuilder();
            bool inQuotes = false;
            int c;

            while ((c = reader.Read()) != -1)
            {
                char ch = (char)c;

                if (ch == '"')
                {
                    if (inQuotes)
                    {
                        int next = reader.Peek();
                        if (next == '"')
                        {
                            reader.Read();
                            currentField.Append('"');
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        inQuotes = true;
                    }
                }
                else if (ch == ',' && !inQuotes)
                {
                    currentRecord.Add(currentField.ToString().Trim());
                    currentField.Clear();
                }
                else if ((ch == '\r' || ch == '\n') && !inQuotes)
                {
                    if (ch == '\r' && reader.Peek() == '\n')
                    {
                        reader.Read();
                    }
                    currentRecord.Add(currentField.ToString().Trim());
                    currentField.Clear();

                    if (currentRecord.Any(f => !string.IsNullOrWhiteSpace(f)))
                    {
                        records.Add(currentRecord);
                    }
                    currentRecord = new List<string>();
                }
                else
                {
                    currentField.Append(ch);
                }
            }

            if (currentField.Length > 0 || currentRecord.Count > 0)
            {
                currentRecord.Add(currentField.ToString().Trim());
                if (currentRecord.Any(f => !string.IsNullOrWhiteSpace(f)))
                {
                    records.Add(currentRecord);
                }
            }

            return records;
        }

        private static string ValueOrDefault(string? val, string fallback)
        {
            return string.IsNullOrWhiteSpace(val) ? fallback : val.Trim();
        }

        public static List<ShippingLabelModel> GetSampleRecords()
        {
            return new List<ShippingLabelModel>
            {
                new ShippingLabelModel
                {
                    AWBNumber = "VL0085889669701",
                    CourierName = "Valmo",
                    PickupDateBadge = "01/11",
                    HubCode = "UFC-R0",
                    RoutingLine1 = "W1/S11S",
                    RoutingLine2 = "E1/BGUS",
                    RoutingLine3 = "D-1/TQS",
                    PaymentStatus = "Check the payable amount on the app",
                    QRCodeData = "VL0085889669701",
                    CustomerName = "Mamta Kumari",
                    CustomerAddressLine1 = "badi manjhgay",
                    CustomerAddressLine2 = "raipura road",
                    CustomerCity = "badi manjhgay",
                    CustomerState = "Bihar",
                    CustomerPincode = "811213",
                    SKU = "JHSILVERCHATTAR_PINK_EAR",
                    Size = "Free Size",
                    Qty = 1,
                    Color = "Pink",
                    OrderNo = "345497888464241728_1",
                    PurchaseOrderNo = "345497888464241728",
                    InvoiceNo = "ynvgm27384",
                    OrderDate = "24.10.2026",
                    InvoiceDate = "25.10.2026",
                    ItemDescription = "Oxidised Silver Pink Beaded Jhumka Earrings for Women & Girls | Ethnic Drop Earrings - Free Size",
                    ItemHSN = "711790",
                    ItemQty = 1,
                    GrossAmount = 82.77m,
                    Discount = 0.00m,
                    TaxableValue = 80.36m,
                    TaxRateDescription = "IGST @3.0%",
                    TaxAmount = 2.41m,
                    ItemTotal = 82.77m,
                    TotalTaxes = 2.41m,
                    GrandTotal = 82.77m,
                    PlaceOfSupply = "Bihar",
                    SoldBy = "AALOKHAJEWELS",
                    ReturnCompanyName = "AALOKHAJEWELS",
                    ReturnAddressLine1 = "B-69, shantiniketan soc, opp sagar soc ,",
                    ReturnAddressLine2 = "behind gujrat mobile",
                    ReturnArea = "kapodra",
                    ReturnCityStatePin = "surat, Gujarat, 395006",
                    SellerAddress = "B-69, shantiniketan soc, opp sagar soc , behind gujrat mobile, kapodra, surat, Gujarat, 395006",
                    SellerGSTIN = "24BFLPR4388M1ZM"
                },
                new ShippingLabelModel
                {
                    AWBNumber = "VL0085889669700",
                    CourierName = "Valmo",
                    PickupDateBadge = "01/11",
                    HubCode = "UFC-R0",
                    RoutingLine1 = "W1/S11S",
                    RoutingLine2 = "N2/EUS",
                    RoutingLine3 = "A4/DSQ",
                    PaymentStatus = "Check the payable amount on the app",
                    QRCodeData = "VL0085889669700",
                    CustomerName = "Rajneesh Verma",
                    CustomerAddressLine1 = "dilawerpur duriya road",
                    CustomerAddressLine2 = "duriya road dilawerpur, pal medical store",
                    CustomerCity = "puranpur",
                    CustomerState = "Uttar Pradesh",
                    CustomerPincode = "262122",
                    SKU = "MANGG_2",
                    Size = "Free Size",
                    Qty = 1,
                    Color = "NA",
                    OrderNo = "345601163607386112_1",
                    PurchaseOrderNo = "345601163607386112",
                    InvoiceNo = "ynvgm27387",
                    OrderDate = "24.10.2026",
                    InvoiceDate = "25.10.2026",
                    ItemDescription = "Elegant Black Beaded Mangalsutra Necklace for Women | Gold Plated Pendant Jewellery - Free Size",
                    ItemHSN = "711790",
                    ItemQty = 1,
                    GrossAmount = 75.04m,
                    Discount = 0.00m,
                    TaxableValue = 72.85m,
                    TaxRateDescription = "IGST @3.0%",
                    TaxAmount = 2.19m,
                    ItemTotal = 75.04m,
                    TotalTaxes = 2.45m,
                    GrandTotal = 84.04m,
                    PlaceOfSupply = "Uttar Pradesh",
                    SoldBy = "AALOKHAJEWELS",
                    ReturnCompanyName = "AALOKHAJEWELS",
                    ReturnAddressLine1 = "B-69, shantiniketan soc, opp sagar soc ,",
                    ReturnAddressLine2 = "behind gujrat mobile",
                    ReturnArea = "kapodra",
                    ReturnCityStatePin = "surat, Gujarat, 395006",
                    SellerAddress = "B-69, shantiniketan soc, opp sagar soc , behind gujrat mobile, kapodra, surat, Gujarat, 395006",
                    SellerGSTIN = "24BFLPR4388M1ZM"
                },
                new ShippingLabelModel
                {
                    AWBNumber = "VL0085889669698",
                    CourierName = "Valmo",
                    PickupDateBadge = "02/11",
                    HubCode = "UFC-R0",
                    RoutingLine1 = "W1/S11S",
                    RoutingLine2 = "E2/RWS",
                    RoutingLine3 = "G-2/MED",
                    PaymentStatus = "Check the payable amount on the app",
                    QRCodeData = "VL0085889669698",
                    CustomerName = "jakiran",
                    CustomerAddressLine1 = "S.R.D Road hayat medical ke gali main Lalgarh",
                    CustomerAddressLine2 = "Madhupur, Deoghar District",
                    CustomerCity = "Madhupur",
                    CustomerState = "Jharkhand",
                    CustomerPincode = "815353",
                    SKU = "MANGG_4",
                    Size = "Free Size",
                    Qty = 1,
                    Color = "White",
                    OrderNo = "345446729293562496_1",
                    PurchaseOrderNo = "345446729293562496",
                    InvoiceNo = "ynvgm27381",
                    OrderDate = "24.10.2026",
                    InvoiceDate = "25.10.2026",
                    ItemDescription = "Elegant Black Beaded Mangalsutra Necklace for Women | Gold Plated Pendant Jewellery - Free Size",
                    ItemHSN = "711719",
                    ItemQty = 1,
                    GrossAmount = 83.04m,
                    Discount = 0.00m,
                    TaxableValue = 80.62m,
                    TaxRateDescription = "IGST @3.0%",
                    TaxAmount = 2.42m,
                    ItemTotal = 83.04m,
                    TotalTaxes = 2.42m,
                    GrandTotal = 83.04m,
                    PlaceOfSupply = "Jharkhand",
                    SoldBy = "AALOKHAJEWELS",
                    ReturnCompanyName = "AALOKHAJEWELS",
                    ReturnAddressLine1 = "B-69, shantiniketan soc, opp sagar soc ,",
                    ReturnAddressLine2 = "behind gujrat mobile",
                    ReturnArea = "kapodra",
                    ReturnCityStatePin = "surat, Gujarat, 395006",
                    SellerAddress = "B-69, shantiniketan soc, opp sagar soc , behind gujrat mobile, kapodra, surat, Gujarat, 395006",
                    SellerGSTIN = "24BFLPR4388M1ZM"
                },
                new ShippingLabelModel
                {
                    AWBNumber = "VL0085889669702",
                    CourierName = "Valmo",
                    PickupDateBadge = "01/11",
                    HubCode = "UFC-R0",
                    RoutingLine1 = "W1/S11S",
                    RoutingLine2 = "N2/AG2S",
                    RoutingLine3 = "A6/KR6",
                    PaymentStatus = "Check the payable amount on the app",
                    QRCodeData = "VL0085889669702",
                    CustomerName = "Rishabh Kumar",
                    CustomerAddressLine1 = "Barnahal gupta mohla, Bhalu Pradhan bali gali",
                    CustomerAddressLine2 = "Vekuti Devi public school",
                    CustomerCity = "manipuri",
                    CustomerState = "Uttar Pradesh",
                    CustomerPincode = "205261",
                    SKU = "MANG_6",
                    Size = "Free Size",
                    Qty = 1,
                    Color = "Golden",
                    OrderNo = "345479017633218560_1",
                    PurchaseOrderNo = "345479017633218560",
                    InvoiceNo = "ynvgm27389",
                    OrderDate = "24.10.2026",
                    InvoiceDate = "25.10.2026",
                    ItemDescription = "Traditional Gold Plated Mangalsutra with Matching Earrings Set - Free Size",
                    ItemHSN = "711719",
                    ItemQty = 1,
                    GrossAmount = 79.50m,
                    Discount = 0.00m,
                    TaxableValue = 77.18m,
                    TaxRateDescription = "IGST @3.0%",
                    TaxAmount = 2.32m,
                    ItemTotal = 79.50m,
                    TotalTaxes = 2.32m,
                    GrandTotal = 79.50m,
                    PlaceOfSupply = "Uttar Pradesh",
                    SoldBy = "AALOKHAJEWELS",
                    ReturnCompanyName = "AALOKHAJEWELS",
                    ReturnAddressLine1 = "B-69, shantiniketan soc, opp sagar soc ,",
                    ReturnAddressLine2 = "behind gujrat mobile",
                    ReturnArea = "kapodra",
                    ReturnCityStatePin = "surat, Gujarat, 395006",
                    SellerAddress = "B-69, shantiniketan soc, opp sagar soc , behind gujrat mobile, kapodra, surat, Gujarat, 395006",
                    SellerGSTIN = "24BFLPR4388M1ZM"
                }
            };
        }
    }
}

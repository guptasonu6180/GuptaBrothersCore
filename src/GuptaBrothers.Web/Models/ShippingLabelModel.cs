using System;
using System.Collections.Generic;

namespace GuptaBrothers.Web.Models
{
    public class ShippingLabelModel
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        // Courier & Routing details
        public string PaymentStatus { get; set; } = "Check the payable amount on the app"; // e.g., "Check the payable amount on the app" or "Prepaid: Do not collect cash"
        public string CourierName { get; set; } = "Valmo";
        public string PickupBadgeText { get; set; } = "Pickup";
        public string PickupDateBadge { get; set; } = "01/11"; // e.g. "01/11", "31/10", "04/11"
        public string HubCode { get; set; } = "UFC-R0";
        public string RoutingLine1 { get; set; } = "W1/S11S";
        public string RoutingLine2 { get; set; } = "E1/BGUS";
        public string RoutingLine3 { get; set; } = "D-1/TQS";
        public string AWBNumber { get; set; } = "VL0085889669701";
        public string QRCodeData { get; set; } = "";

        // Delivery / Customer Address
        public string CustomerName { get; set; } = "";
        public string CustomerAddressLine1 { get; set; } = "";
        public string CustomerAddressLine2 { get; set; } = "";
        public string CustomerCity { get; set; } = "";
        public string CustomerState { get; set; } = "";
        public string CustomerPincode { get; set; } = "";
        public string CustomerPhone { get; set; } = "";

        // Return Address
        public string ReturnCompanyName { get; set; } = "Gupta Brothers";
        public string ReturnAddressLine1 { get; set; } = "B-69, shantiniketan soc, opp sagar soc ,";
        public string ReturnAddressLine2 { get; set; } = "behind gujrat mobile";
        public string ReturnArea { get; set; } = "kapodra";
        public string ReturnCityStatePin { get; set; } = "surat, Gujarat, 395006";

        // Product Details
        public string SKU { get; set; } = "";
        public string Size { get; set; } = "Free Size";
        public int Qty { get; set; } = 1;
        public string Color { get; set; } = "NA";
        public string OrderNo { get; set; } = "";

        // Tax Invoice Header
        public string BillToShipTo { get; set; } = "";
        public string PlaceOfSupply { get; set; } = "";
        public string SoldBy { get; set; } = "Gupta Brothers";
        public string SellerAddress { get; set; } = "Gupta Brothers, D-502 FIFTH FLOOR Angan Residency Puna Simada Road , Surat, Gujarat, 395010";
        public string SellerGSTIN { get; set; } = "24BFLPR4388M1ZM";
        public string PurchaseOrderNo { get; set; } = "";
        public string InvoiceNo { get; set; } = "";
        public string OrderDate { get; set; } = DateTime.Now.ToString("dd.MM.yyyy");
        public string InvoiceDate { get; set; } = DateTime.Now.ToString("dd.MM.yyyy");

        // Item details
        public string ItemDescription { get; set; } = "";
        public string ItemHSN { get; set; } = "711790";
        public int ItemQty { get; set; } = 1;
        public decimal GrossAmount { get; set; } = 0.00m;
        public decimal Discount { get; set; } = 0.00m;
        public decimal TaxableValue { get; set; } = 0.00m;
        public string TaxRateDescription { get; set; } = "IGST @3.0%";
        public decimal TaxAmount { get; set; } = 0.00m;
        public decimal ItemTotal { get; set; } = 0.00m;

        // Convenient Aliases
        public string HSNCode
        {
            get => ItemHSN;
            set => ItemHSN = value;
        }

        public decimal IGSTAmount
        {
            get => TaxAmount;
            set => TaxAmount = value;
        }

        // Other charges (Shipping / Logistics)
        public string OtherChargesHSN { get; set; } = "711790";
        public string OtherChargesQty { get; set; } = "NA";
        public decimal OtherChargesGross { get; set; } = 0.00m;
        public decimal OtherChargesDiscount { get; set; } = 0.00m;
        public decimal OtherChargesTaxable { get; set; } = 0.00m;
        public string OtherChargesTaxRate { get; set; } = "IGST @3.0%";
        public decimal OtherChargesTax { get; set; } = 0.00m;
        public decimal OtherChargesTotal { get; set; } = 0.00m;

        public decimal OtherCharges
        {
            get => OtherChargesGross;
            set => OtherChargesGross = value;
        }

        public decimal OtherChargesIGST
        {
            get => OtherChargesTax;
            set => OtherChargesTax = value;
        }

        // Totals
        public decimal TotalTaxes { get; set; } = 0.00m;
        public decimal GrandTotal { get; set; } = 0.00m;

        // In Words
        public string AmountInWords
        {
            get
            {
                long val = (long)Math.Round(GrandTotal);
                if (val <= 0) return "Zero Rupees Only";
                return ConvertNumberToWords(val) + " Rupees Only";
            }
        }

        // Additional / Computed
        public DateTime CreatedAt { get; set; } = DateTime.Now;

        // Dynamic QR and Barcode representations
        public string? QrCodeSvg { get; set; }
        public string? QrCodeBase64 { get; set; }
        public string? BarcodeSvg { get; set; }

        public void AutoCalculate()
        {
            if (string.IsNullOrWhiteSpace(PurchaseOrderNo) && !string.IsNullOrWhiteSpace(OrderNo))
            {
                var idx = OrderNo.LastIndexOf('_');
                PurchaseOrderNo = idx > 0 ? OrderNo.Substring(0, idx) : OrderNo;
            }

            if (ItemQty <= 0) ItemQty = Qty > 0 ? Qty : 1;

            if (TaxableValue == 0 && GrossAmount > 0)
            {
                var net = GrossAmount - Discount;
                TaxableValue = Math.Round(net / 1.03m, 2);
                TaxAmount = Math.Round(net - TaxableValue, 2);
                ItemTotal = net;
            }

            if (OtherChargesGross > 0 && OtherChargesTaxable == 0)
            {
                OtherChargesTaxable = Math.Round(OtherChargesGross / 1.03m, 2);
                OtherChargesTax = Math.Round(OtherChargesGross - OtherChargesTaxable, 2);
                OtherChargesTotal = OtherChargesGross;
            }

            TotalTaxes = TaxAmount + OtherChargesTax;
            GrandTotal = ItemTotal + OtherChargesTotal;

            if (string.IsNullOrWhiteSpace(BillToShipTo))
            {
                var parts = new List<string>();
                if (!string.IsNullOrWhiteSpace(CustomerName)) parts.Add(CustomerName);
                if (!string.IsNullOrWhiteSpace(CustomerAddressLine1)) parts.Add(CustomerAddressLine1);
                if (!string.IsNullOrWhiteSpace(CustomerAddressLine2)) parts.Add(CustomerAddressLine2);
                if (!string.IsNullOrWhiteSpace(CustomerCity)) parts.Add(CustomerCity);
                if (!string.IsNullOrWhiteSpace(CustomerState)) parts.Add(CustomerState);
                if (!string.IsNullOrWhiteSpace(CustomerPincode)) parts.Add(CustomerPincode);
                if (!string.IsNullOrWhiteSpace(PlaceOfSupply)) parts.Add($"Place of Supply: {PlaceOfSupply}");
                else if (!string.IsNullOrWhiteSpace(CustomerState)) parts.Add($"Place of Supply: {CustomerState}");

                BillToShipTo = string.Join(", ", parts);
            }

            if (string.IsNullOrWhiteSpace(PlaceOfSupply) && !string.IsNullOrWhiteSpace(CustomerState))
            {
                PlaceOfSupply = CustomerState;
            }
        }

        private static string ConvertNumberToWords(long number)
        {
            if (number == 0) return "Zero";
            if (number < 0) return "Minus " + ConvertNumberToWords(Math.Abs(number));

            string words = "";

            if ((number / 10000000) > 0)
            {
                words += ConvertNumberToWords(number / 10000000) + " Crore ";
                number %= 10000000;
            }

            if ((number / 100000) > 0)
            {
                words += ConvertNumberToWords(number / 100000) + " Lakh ";
                number %= 100000;
            }

            if ((number / 1000) > 0)
            {
                words += ConvertNumberToWords(number / 1000) + " Thousand ";
                number %= 1000;
            }

            if ((number / 100) > 0)
            {
                words += ConvertNumberToWords(number / 100) + " Hundred ";
                number %= 100;
            }

            if (number > 0)
            {
                var unitsMap = new[] { "Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
                var tensMap = new[] { "Zero", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

                if (number < 20)
                    words += unitsMap[number];
                else
                {
                    words += tensMap[number / 10];
                    if ((number % 10) > 0)
                        words += " " + unitsMap[number % 10];
                }
            }

            return words.Trim();
        }
    }

    public class LabelSettingsModel
    {
        public string DefaultReturnCompanyName { get; set; } = "Gupta Brothers";
        public string DefaultReturnAddressLine1 { get; set; } = "B-69, shantiniketan soc, opp sagar soc ,";
        public string DefaultReturnAddressLine2 { get; set; } = "behind gujrat mobile";
        public string DefaultReturnArea { get; set; } = "kapodra";
        public string DefaultReturnCityStatePin { get; set; } = "surat, Gujarat, 395006";

        public string DefaultSoldBy { get; set; } = "Gupta Brothers";
        public string DefaultSellerAddress { get; set; } = "Gupta Brothers, D-502 FIFTH FLOOR Angan Residency Puna Simada Road , Surat, Gujarat, 395010";
        public string DefaultSellerGSTIN { get; set; } = "24BFLPR4388M1ZM";

        public string DefaultCourierName { get; set; } = "Valmo";
        public string DefaultHubCode { get; set; } = "UFC-R0";
        public string QrCodeContentPattern { get; set; } = "{AWB}"; // Options: {AWB}, {AWB}|{OrderNo}, {OrderNo}
        public decimal DefaultGstRate { get; set; } = 3.0m;
    }
}

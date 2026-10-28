using System;
using System.Collections.Generic;
using System.Text;
using QRCoder;

namespace GuptaBrothers.Web.Services
{
    public class QrCodeBarcodeService : IQrCodeBarcodeService
    {
        public string GenerateQrCodePngBase64(string content, int pixelsPerModule = 10)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                content = "NO_DATA";
            }

            try
            {
                using var qrGenerator = new QRCodeGenerator();
                var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
                var pngQrCode = new PngByteQRCode(qrCodeData);
                byte[] bytes = pngQrCode.GetGraphic(pixelsPerModule);
                return "data:image/png;base64," + Convert.ToBase64String(bytes);
            }
            catch
            {
                return "";
            }
        }

        public string GenerateQrCodeSvg(string content, int pixelsPerModule = 4)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                content = "NO_DATA";
            }

            try
            {
                using var qrGenerator = new QRCodeGenerator();
                var qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.M);
                var qrCode = new SvgQRCode(qrCodeData);
                return qrCode.GetGraphic(pixelsPerModule);
            }
            catch (Exception ex)
            {
                return $"<!-- QR Error: {ex.Message} -->";
            }
        }

        public string GenerateBarcodeSvg(string text, int height = 65, int barWidth = 3)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                text = "NO_DATA";
            }

            try
            {
                return Code128SvgGenerator.Generate(text.Trim(), height, barWidth);
            }
            catch (Exception ex)
            {
                return $"<!-- Barcode Error: {ex.Message} -->";
            }
        }
    }

    /// <summary>
    /// Pure C# Code 128 (Type B) vector SVG generator.
    /// Scalable SVG barcode with zero native dependencies.
    /// </summary>
    public static class Code128SvgGenerator
    {
        private static readonly string[] Patterns = new string[]
        {
            "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213", // 0-9
            "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132", // 10-19
            "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211", // 20-29
            "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313", // 30-39
            "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331", // 40-49
            "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111", // 50-59
            "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214", // 60-69
            "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111", // 70-79
            "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141", // 80-89
            "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141", // 90-99
            "114131", "311141", "411131", "211412", "211214", "211232", "2331112" // 100-106
        };

        public static string Generate(string text, int height = 40, int barWidth = 2)
        {
            var codes = new List<int>();
            codes.Add(104); // Start Code B
            long checksum = 104;

            for (int i = 0; i < text.Length; i++)
            {
                int val = text[i] - 32;
                if (val < 0 || val > 95) val = 0;
                codes.Add(val);
                checksum += (long)(i + 1) * val;
            }

            codes.Add((int)(checksum % 103));
            codes.Add(106); // Stop code

            var patternBuilder = new StringBuilder();
            foreach (var code in codes)
            {
                patternBuilder.Append(Patterns[code]);
            }

            string combined = patternBuilder.ToString();
            int totalModules = 0;
            foreach (char c in combined)
            {
                totalModules += (c - '0');
            }

            int svgWidth = totalModules * barWidth;
            var svg = new StringBuilder();
            svg.Append($"<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 {svgWidth} {height}\" width=\"100%\" height=\"100%\" preserveAspectRatio=\"none\" style=\"display:block;max-width:100%;max-height:100%;\">");
            svg.Append($"<rect width=\"100%\" height=\"100%\" fill=\"#ffffff\"/>");

            int x = 0;
            bool isBar = true;
            foreach (char c in combined)
            {
                int widthInModules = (c - '0');
                int pixelWidth = widthInModules * barWidth;
                if (isBar)
                {
                    svg.Append($"<rect x=\"{x}\" y=\"0\" width=\"{pixelWidth}\" height=\"{height}\" fill=\"#000000\"/>");
                }
                x += pixelWidth;
                isBar = !isBar;
            }

            svg.Append("</svg>");
            return svg.ToString();
        }
    }
}

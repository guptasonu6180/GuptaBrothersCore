using System;

namespace GuptaBrothers.Web.Services
{
    public interface IQrCodeBarcodeService
    {
        string GenerateQrCodePngBase64(string content, int pixelsPerModule = 10);
        string GenerateQrCodeSvg(string content, int pixelsPerModule = 4);
        string GenerateBarcodeSvg(string text, int height = 65, int barWidth = 3);
    }
}

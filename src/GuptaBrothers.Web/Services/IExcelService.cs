using System.Collections.Generic;
using System.IO;
using GuptaBrothers.Web.Models;

namespace GuptaBrothers.Web.Services
{
    public interface IExcelService
    {
        byte[] GenerateSampleTemplate();
        List<ShippingLabelModel> ImportFromExcel(Stream stream, LabelSettingsModel? defaultSettings = null);
        List<ShippingLabelModel> ImportFromFile(Stream stream, string fileName, LabelSettingsModel? defaultSettings = null);
        byte[] ExportToExcel(IEnumerable<ShippingLabelModel> labels);
    }
}

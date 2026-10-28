using System.Collections.Generic;
using GuptaBrothers.Web.Models;

namespace GuptaBrothers.Web.Services
{
    public interface IShippingLabelService
    {
        List<ShippingLabelModel> GetAll(string? searchTerm = null);
        ShippingLabelModel? GetById(string id);
        List<ShippingLabelModel> GetByIds(IEnumerable<string> ids);
        void Add(ShippingLabelModel label);
        void AddRange(IEnumerable<ShippingLabelModel> labels);
        void Update(ShippingLabelModel label);
        bool Delete(string id);
        void ClearAll();
        void ResetToDefaultSampleData();
        LabelSettingsModel GetSettings();
        void SaveSettings(LabelSettingsModel settings);
        void EnsureQrAndBarcode(ShippingLabelModel label);
    }
}

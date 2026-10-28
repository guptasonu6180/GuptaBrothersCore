using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.AspNetCore.Hosting;
using GuptaBrothers.Web.Models;

namespace GuptaBrothers.Web.Services
{
    public class ShippingLabelService : IShippingLabelService
    {
        private readonly IQrCodeBarcodeService _qrBarcodeService;
        private readonly IWebHostEnvironment _env;
        private readonly object _lock = new object();
        private List<ShippingLabelModel> _cache = new List<ShippingLabelModel>();
        private LabelSettingsModel _settings = new LabelSettingsModel();
        private readonly string _dataFilePath;
        private readonly string _settingsFilePath;

        public ShippingLabelService(IQrCodeBarcodeService qrBarcodeService, IWebHostEnvironment env)
        {
            _qrBarcodeService = qrBarcodeService;
            _env = env;

            var dataDir = Path.Combine(_env.ContentRootPath, "App_Data");
            if (!Directory.Exists(dataDir))
            {
                Directory.CreateDirectory(dataDir);
            }

            _dataFilePath = Path.Combine(dataDir, "shipping_labels.json");
            _settingsFilePath = Path.Combine(dataDir, "label_settings.json");

            LoadSettings();
            LoadData();
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    var json = File.ReadAllText(_settingsFilePath);
                    var loaded = JsonSerializer.Deserialize<LabelSettingsModel>(json);
                    if (loaded != null)
                    {
                        _settings = loaded;
                        return;
                    }
                }
            }
            catch { }
            _settings = new LabelSettingsModel();
        }

        public void SaveSettings(LabelSettingsModel settings)
        {
            lock (_lock)
            {
                _settings = settings;
                try
                {
                    var json = JsonSerializer.Serialize(_settings, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(_settingsFilePath, json);
                }
                catch { }
            }
        }

        public LabelSettingsModel GetSettings()
        {
            lock (_lock)
            {
                return _settings;
            }
        }

        private void LoadData()
        {
            lock (_lock)
            {
                try
                {
                    if (File.Exists(_dataFilePath))
                    {
                        var json = File.ReadAllText(_dataFilePath);
                        var list = JsonSerializer.Deserialize<List<ShippingLabelModel>>(json);
                        if (list != null && list.Count > 0)
                        {
                            _cache = list;
                            return;
                        }
                    }
                }
                catch { }

                // Seed with realistic default data
                _cache = ExcelService.GetSampleRecords();
                PersistData();
            }
        }

        private void PersistData()
        {
            try
            {
                var json = JsonSerializer.Serialize(_cache, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_dataFilePath, json);
            }
            catch { }
        }

        public List<ShippingLabelModel> GetAll(string? searchTerm = null)
        {
            lock (_lock)
            {
                var query = _cache.AsEnumerable();
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    var term = searchTerm.Trim().ToLowerInvariant();
                    query = query.Where(x =>
                        (x.AWBNumber != null && x.AWBNumber.ToLowerInvariant().Contains(term)) ||
                        (x.CustomerName != null && x.CustomerName.ToLowerInvariant().Contains(term)) ||
                        (x.OrderNo != null && x.OrderNo.ToLowerInvariant().Contains(term)) ||
                        (x.SKU != null && x.SKU.ToLowerInvariant().Contains(term)) ||
                        (x.CustomerCity != null && x.CustomerCity.ToLowerInvariant().Contains(term)) ||
                        (x.CustomerState != null && x.CustomerState.ToLowerInvariant().Contains(term)) ||
                        (x.InvoiceNo != null && x.InvoiceNo.ToLowerInvariant().Contains(term))
                    );
                }

                var list = query.OrderByDescending(x => x.CreatedAt).ToList();
                return list;
            }
        }

        public ShippingLabelModel? GetById(string id)
        {
            lock (_lock)
            {
                var item = _cache.FirstOrDefault(x => x.Id == id || x.AWBNumber == id);
                if (item != null)
                {
                    EnsureQrAndBarcode(item);
                }
                return item;
            }
        }

        public List<ShippingLabelModel> GetByIds(IEnumerable<string> ids)
        {
            var idSet = new HashSet<string>(ids, StringComparer.OrdinalIgnoreCase);
            lock (_lock)
            {
                var items = _cache.Where(x => idSet.Contains(x.Id) || idSet.Contains(x.AWBNumber)).ToList();
                foreach (var item in items)
                {
                    EnsureQrAndBarcode(item);
                }
                return items;
            }
        }

        public void Add(ShippingLabelModel label)
        {
            lock (_lock)
            {
                label.AutoCalculate();
                EnsureQrAndBarcode(label);
                _cache.Insert(0, label);
                PersistData();
            }
        }

        public void AddRange(IEnumerable<ShippingLabelModel> labels)
        {
            var list = labels.ToList();
            Parallel.ForEach(list, l =>
            {
                l.AutoCalculate();
                EnsureQrAndBarcode(l);
            });

            lock (_lock)
            {
                _cache.InsertRange(0, list);
                PersistData();
            }
        }

        public void Update(ShippingLabelModel label)
        {
            lock (_lock)
            {
                var existing = _cache.FirstOrDefault(x => x.Id == label.Id);
                if (existing != null)
                {
                    _cache.Remove(existing);
                    label.AutoCalculate();
                    EnsureQrAndBarcode(label);
                    _cache.Insert(0, label);
                    PersistData();
                }
            }
        }

        public bool Delete(string id)
        {
            lock (_lock)
            {
                var existing = _cache.FirstOrDefault(x => x.Id == id || x.AWBNumber == id);
                if (existing != null)
                {
                    _cache.Remove(existing);
                    PersistData();
                    return true;
                }
                return false;
            }
        }

        public void ClearAll()
        {
            lock (_lock)
            {
                _cache.Clear();
                PersistData();
            }
        }

        public void ResetToDefaultSampleData()
        {
            lock (_lock)
            {
                _cache = ExcelService.GetSampleRecords();
                PersistData();
            }
        }

        public void EnsureQrAndBarcode(ShippingLabelModel label)
        {
            // Barcode
            var barcodeText = !string.IsNullOrWhiteSpace(label.AWBNumber) ? label.AWBNumber.Trim() : (label.OrderNo ?? "").Trim();
            if (string.IsNullOrEmpty(label.BarcodeSvg))
            {
                label.BarcodeSvg = _qrBarcodeService.GenerateBarcodeSvg(barcodeText, height: 65, barWidth: 3);
            }

            // QR Code: ONLY the AWB Number value and nothing more (e.g. "VL0085889669704")
            string qrData = !string.IsNullOrWhiteSpace(label.AWBNumber) ? label.AWBNumber.Trim() : (label.OrderNo ?? "").Trim();
            label.QRCodeData = qrData;

            if (string.IsNullOrEmpty(label.QrCodeBase64))
            {
                label.QrCodeBase64 = _qrBarcodeService.GenerateQrCodePngBase64(qrData, pixelsPerModule: 10);
            }
            if (string.IsNullOrEmpty(label.QrCodeSvg))
            {
                label.QrCodeSvg = _qrBarcodeService.GenerateQrCodeSvg(qrData, pixelsPerModule: 6);
            }
        }
    }
}

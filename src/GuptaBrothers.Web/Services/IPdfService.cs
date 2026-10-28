using System.Collections.Generic;
using GuptaBrothers.Web.Models;

namespace GuptaBrothers.Web.Services
{
    public interface IPdfService
    {
        byte[] GenerateSingleLabelPdf(ShippingLabelModel label);
        byte[] GenerateBatchLabelPdf(IEnumerable<ShippingLabelModel> labels);
    }
}

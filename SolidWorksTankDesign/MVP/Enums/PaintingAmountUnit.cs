using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace SolidWorksTankDesign.MVP.Enums
{
    public enum PaintingAmountUnit
    {
        [Description("%")]
        Percentage,
        [Description("mm")]
        Millimeters
    }
}

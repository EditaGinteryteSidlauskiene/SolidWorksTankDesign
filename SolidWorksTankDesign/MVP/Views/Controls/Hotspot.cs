using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign.MVP.Views.Controls
{
    public class Hotspot
    {
        // Relative (0..1) position so resizing works
        public float X { get; set; }
        public float Y { get; set; }

        public NozzleVerticalReferenceType? ReferenceType { get; set; }

        public NozzleBottomReferencePoint? BottomReferencePoint { get; set; }

        public NozzlePropertiesType? NozzlePropertiesType { get; set; }
    }
}

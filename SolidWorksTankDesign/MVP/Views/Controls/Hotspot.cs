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

        public NozzleTopReferenceType? ReferenceType { get; set; }

        public NozzleBottomReferencePoint? BottomReferencePoint { get; set; }

        public NozzlePropertiesType? NozzlePropertiesType { get; set; }

        public NozzleReferenceType? NozzleReferenceType { get; set; }

        public bool IsNozzleLength { get; set; }

        public FlipDot? FlipDot { get; set; }

        public bool IsRotationArrow { get; set; }

        public float Tolerance { get; set; } = 6f;

        public Action OnClick { get; set; }
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign.MVP.Models
{
    public interface INozzleModel
    {
        void AddNozzle(string projectFolderPath, NozzleReferenceType referenceType, double distance);

        void RepositionNozzle(bool isOffsetPositive, double distance, bool isRotationDirectionPositive, double angle);
    }
}

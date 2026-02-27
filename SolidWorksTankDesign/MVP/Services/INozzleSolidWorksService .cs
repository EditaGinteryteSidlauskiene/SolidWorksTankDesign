using System;
using System.Collections.Generic;
using System.Linq;
namespace SolidWorksTankDesign.MVP.Services
{
    public interface INozzleSolidWorksService
    {
        void AddNozzle(string projectFolderPath, string selectedNozzleRef, double distance);

        void RepositionNozzle(
            bool isOffsetPositive,
            double distance,
            bool isRotationDirectionPositive,
            double angle);
    }
}

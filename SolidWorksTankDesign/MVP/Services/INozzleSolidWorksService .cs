using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
namespace SolidWorksTankDesign.MVP.Services
{
    public interface INozzleSolidWorksService
    {
        void AddNozzle(Guid compartmentConfigId, NozzleReferenceType referenceType, double distance);

        void ApplyNozzleChanges();
    }
}

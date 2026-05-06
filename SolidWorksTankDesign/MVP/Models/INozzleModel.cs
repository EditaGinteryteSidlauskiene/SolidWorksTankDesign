using SolidWorksTankDesign.MVP.Enums;
using System;

namespace SolidWorksTankDesign.MVP.Models
{
    public interface INozzleModel
    {
        void AddNozzle(Guid compartmentConfigId, NozzleReferenceType referenceType, double distance);
        void RepositionNozzle(bool isOffsetPositive, double distance, bool isRotationDirectionPositive, double angle);
        void ApplyNozzleChanges();
    }
}
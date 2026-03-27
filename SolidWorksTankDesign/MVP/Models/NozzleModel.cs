using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Services;
using System;

namespace SolidWorksTankDesign.MVP.Models
{
    public class NozzleModel : INozzleModel
    {
        private readonly INozzleSolidWorksService _nozzleSWService;

        // Service injection - if present, model delegates SolidWorks work to the service.
        public NozzleModel(INozzleSolidWorksService swService)
        {
            _nozzleSWService = swService ?? throw new ArgumentNullException(nameof(swService));
        }

        // INozzleModel implementation - delegate to service when available.
        public void AddNozzle(Guid compartmentConfigId, NozzleReferenceType referenceType, double distance)
        {
            // Basic argument validation
            if (!Enum.IsDefined(typeof(NozzleReferenceType), referenceType))
                throw new ArgumentException("Invalid nozzle reference type.", nameof(referenceType));

            if (double.IsNaN(distance) || double.IsInfinity(distance) || distance < 0)
                throw new ArgumentOutOfRangeException(nameof(distance), "Distance must be a non-negative finite value (meters).");

            if (_nozzleSWService == null) throw new InvalidOperationException("NozzleSolidWorksService not provided.");

            _nozzleSWService.AddNozzle(compartmentConfigId, referenceType, distance);
        }

        public void RepositionNozzle(bool isOffsetPositive, double distance, bool isRotationDirectionPositive, double angle)
        {
            if (double.IsNaN(distance) || double.IsInfinity(distance))
                throw new ArgumentOutOfRangeException(nameof(distance), "Distance must be a finite value (meters).");

            if (double.IsNaN(angle) || double.IsInfinity(angle) || angle < 0)
                throw new ArgumentOutOfRangeException(nameof(angle), "Angle must be a non-negative finite value.");

            if (_nozzleSWService == null) throw new InvalidOperationException("NozzleSolidWorksService not provided.");

            _nozzleSWService.RepositionNozzle(isOffsetPositive, distance, isRotationDirectionPositive, angle);
        }
    }
}

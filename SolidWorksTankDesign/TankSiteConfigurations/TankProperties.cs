using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using System.Collections.Generic;

namespace SolidWorksTankDesign.TankSiteConfigurations
{
    public class TankProperties
    {
        public string SerialNumber { get; set; }

        public ConstructionStandard ConstructionStandard { get; set; }

        public double NominalDiameter { get; set; }

        public Class Class {  get; set; }

        public TankType Type { get; set; }

        public double MinOperatingTemperature { get; set; }

        public double MaxOperatingTemperature { get; set;}

        public double MaxOperatingPressure {  get; set; }

        public double ShellLeakTestPressure { get; set; }

        public double InterstitialSpaceLeakTestPressure { get; set; }

        public string InterstitialSpace {  get; set; }

        public string LeakDetectionSystem { get; set; }

        public Treatment ExternalSurfaceTreatment { get; set; }

        public List<CompartmentConfiguration> CompartmentsConfigurations { get; set; } = new List<CompartmentConfiguration>();
    }
}

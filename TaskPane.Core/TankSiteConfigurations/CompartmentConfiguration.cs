using SolidWorksTankDesign.Treatments;
using System.Collections.Generic;
using System.Drawing;

namespace SolidWorksTankDesign.TankSiteConfigurations
{
    internal class CompartmentConfiguration
    {
        public double Volume { get; set; }
        public double Length { get; set; }
        public Treatment InternalSurfaceTreatment { get; set; }
        public double Amount { get; set; }
        public string AmountUnits { get; set; }
        public DishedEndAlignment LeftDishedEndAlignment { get; set; }

        public List<Nozzle> Nozzles { get; set; } = new List<Nozzle>();

        public string LeftEndConnection { get; set; }
    }
}

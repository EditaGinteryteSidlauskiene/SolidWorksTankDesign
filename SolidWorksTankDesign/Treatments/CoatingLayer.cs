using System;

namespace SolidWorksTankDesign.Treatments
{
    public class CoatingLayer : ICloneable
    {
        public string ProductName { get; set; }
        public double DryFilmThickness { get; set; }
        public double WetFilmThickness { get; set; }
        public double FactualConsumption { get; set; }
        public string Units { get; set; }

        public object Clone()
        {
            return new CoatingLayer
            {
                ProductName = this.ProductName,
                DryFilmThickness = this.DryFilmThickness,
                WetFilmThickness = this.WetFilmThickness,
                FactualConsumption = this.FactualConsumption,
                Units = this.Units
            };
        }
    }
}

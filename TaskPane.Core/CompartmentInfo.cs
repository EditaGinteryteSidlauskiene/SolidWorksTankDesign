using SolidWorksTankDesign.Treatments;

namespace TaskPane.Core
{
    public class CompartmentInfo
    {
        public enum Connection
        {
            None,
            Closed,
            Open
        }
        public string Title { get; set; }

        public double Volume { get; set; }

        public double Length { get; set; }

        public Treatment internalSurfaceTreatment { get; set; }

        public double PaintingAmount { get; set; }

        public string PaintingUnits { get; set; }

        public string LeftDishedEndAlignment { get; set; }

        public Connection ConnectionWithPreviousCompartment { get; set; }

        public string RightDishedEndAlignment { get; set;}

        public Connection ConnectionWithNextCompartment { get; set; }
    }
}

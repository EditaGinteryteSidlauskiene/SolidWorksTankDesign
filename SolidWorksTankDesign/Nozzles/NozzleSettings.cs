using Newtonsoft.Json;
using System.Linq;

namespace SolidWorksTankDesign
{
    public class NozzleSettings
    {
        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; set; }

        [JsonProperty("PIDNozzleAxis")]
        public byte[] PIDNozzleAxis { get; set; }

        [JsonProperty("PIDPositionPlane")]
        public byte[] PIDPositionPlane { get; set; }

        [JsonProperty("PIDExternalPoint")]
        public byte[] PIDExternalPoint { get; set; }

        [JsonProperty("PIDInternalPoint")]
        public byte[] PIDInternalPoint { get; set; }

        [JsonProperty("PIDInsidePoint")]
        public byte[] PIDInsidePoint { get; set; }

        [JsonProperty("PIDMidPoint")]
        public byte[] PIDMidPoint { get; set; }

        [JsonProperty("PIDNozzleRightRefPlane")]
        public byte[] PIDNozzleRightRefPlane { get; set; }

        //--------------- RENAME ----------------------
        [JsonProperty("PIDCutPlane")]
        public byte[] PIDCutPlane { get; set; }

        [JsonProperty("PIDSketch")]
        public byte[] PIDSketch { get; set; }

        [JsonProperty("PIDComponent")]
        public byte[] PIDComponent { get; set; }

        [JsonProperty("PIDPositionPlaneMate")]
        public byte[] PIDPositionPlaneMate { get; set; }

        [JsonProperty("PIDTopPlaneMate")]
        public byte[] PIDTopPlaneMate { get; set; }

        [JsonProperty("PIDNozzleAssemblyComp")]
        public byte[] PIDNozzleAssemblyComp { get; set; }

        [JsonProperty("PIDCutOutPlane")]
        public byte[] PIDCutOutPlane { get; set; }

        [JsonProperty("PIDCutExtrude")]
        public byte[] PIDCutExtrude { get; set; }

        public NozzleSettings() { }

        public NozzleSettings DeepClone()
        {
            return new NozzleSettings
            {
                PIDCenterAxis = this.PIDCenterAxis?.ToArray(),
                PIDNozzleAxis = this.PIDNozzleAxis?.ToArray(),
                PIDPositionPlane = this.PIDPositionPlane?.ToArray(),
                PIDExternalPoint = this.PIDExternalPoint?.ToArray(),
                PIDInternalPoint = this.PIDInternalPoint?.ToArray(),
                PIDInsidePoint = this.PIDInsidePoint?.ToArray(),
                PIDMidPoint = this.PIDMidPoint?.ToArray(),
                PIDNozzleRightRefPlane = this.PIDNozzleRightRefPlane?.ToArray(),
                PIDCutPlane = this.PIDCutPlane?.ToArray(),
                PIDSketch = this.PIDSketch?.ToArray(),
                PIDComponent = this.PIDComponent?.ToArray(),
                PIDPositionPlaneMate = this.PIDPositionPlaneMate?.ToArray(),
                PIDTopPlaneMate = this.PIDTopPlaneMate?.ToArray(),
                PIDNozzleAssemblyComp = this.PIDNozzleAssemblyComp?.ToArray(),
                PIDCutOutPlane = this.PIDCutOutPlane?.ToArray(),
                PIDCutExtrude = this.PIDCutExtrude?.ToArray()
            };
        }
    }
}

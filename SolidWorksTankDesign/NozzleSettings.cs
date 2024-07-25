using Newtonsoft.Json;

namespace SolidWorksTankDesign
{
    internal class NozzleSettings
    {
        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; set; }

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

        public NozzleSettings() { }
    }
}

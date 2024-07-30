using Newtonsoft.Json;

namespace SolidWorksTankDesign
{
    internal class CompartmentSettings
    {
        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; set; }

        [JsonProperty("PIDComponent")]
        public byte[] PIDComponent { get; set; }

        [JsonProperty("PIDLeftEndPlane")]
        public byte[] PIDLeftEndPlane { get; set; }

        [JsonProperty("PIDRightEndPlane")]
        public byte[] PIDRightEndPlane { get; set; }

        [JsonProperty("PIDLeftEndMate")]
        public byte[] PIDLeftEndMate { get; set; }

        [JsonProperty("PIDFrontPlaneMate")]
        public byte[] PIDFrontPlaneMate { get; set; }

        [JsonProperty("PIDCenterAxisMate")]
        public byte[] PIDCenterAxisMate { get; set; }

        [JsonProperty("PIDDishedEndPositionPlane")]
        public byte[] PIDDishedEndPositionPlane { get; set; }

        public CompartmentSettings() { }
    }
}

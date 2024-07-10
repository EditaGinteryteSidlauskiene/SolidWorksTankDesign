using Newtonsoft.Json;

namespace SolidWorksTankDesign
{
    internal class CylindricalShellSettings
    {
        [JsonProperty("PIDLeftEndPlane")]
        public byte[] PIDLeftEndPlane { get; set; }

        [JsonProperty("PIDRightEndPlane")]
        public byte[] PIDRightEndPlane { get; set; }

        [JsonProperty("PIDComponent")]
        public byte[] PIDComponent { get; set; }

        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; set; }

        [JsonProperty("PIDCenterAxisMate")]
        public byte[] PIDCenterAxisMate { get; set; }

        [JsonProperty("PIDLeftEndMate")]
        public byte[] PIDLeftEndMate { get; set; }

        [JsonProperty("PIDFrontPlaneMate")]
        public byte[] PIDFrontPlaneMate { get; set; }

        public CylindricalShellSettings() { }
    }
}

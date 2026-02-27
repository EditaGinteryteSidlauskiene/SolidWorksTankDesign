using Newtonsoft.Json;
using System;

namespace SolidWorksTankDesign
{
    public class CompartmentSettings
    {
        [JsonProperty("ID")]
        public Guid ID { get; set; }

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

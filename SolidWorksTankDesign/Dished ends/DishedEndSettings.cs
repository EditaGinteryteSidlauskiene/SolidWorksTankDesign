﻿using Newtonsoft.Json;

namespace SolidWorksTankDesign
{
    internal class DishedEndSettings
    {
        [JsonProperty("PIDPositionPlane")]
        public  byte[] PIDPositionPlane { get; set; }

        [JsonProperty("PIDComponent")]
        public byte[] PIDComponent { get; set; }

        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; set; }

        [JsonProperty("PIDCenterAxisMate")]
        public byte[] PIDCenterAxisMate { get; set; }

        [JsonProperty("PIDRightPlaneMate")]
        public byte[] PIDRightPlaneMate { get; set; }

        [JsonProperty("PIDFrontPlaneMate")]
        public byte[] PIDFrontPlaneMate { get; set; }

        public DishedEndSettings() { }


    }
}

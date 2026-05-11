using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SolidWorksTankDesign
{
    public class NozzleSettings
    {
        [JsonProperty("ID")]
        public Guid ID { get; set; }

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

        [JsonProperty("PIDTopPoint")]
        public byte[] PIDTopPoint { get; set; }

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

        [JsonProperty("PIDCenterlineWallIntersection")]
        public byte[] PIDCenterlineWallIntersection { get; set; }

        [JsonProperty("NozzleAssemblyComponents")]
        public List<NozzleAssemblyComponent> NozzleAssemblyComponents { get; set; } = new List<NozzleAssemblyComponent>();

        [JsonProperty("RotationAngleDegrees")]
        public double RotationAngleDegrees { get; set; }

        [JsonProperty("OffsetMeters")]
        public double OffsetMeters { get; set; }

        [JsonProperty("OffsetPosition")]
        public FlipDot OffsetPosition { get; set; } = FlipDot.Central;

        [JsonProperty("BottomReferencePoint")]
        public NozzleBottomReferencePoint BottomReferencePoint { get; set; }

        [JsonProperty("DistanceFromBottomReferenceMeters")]
        public double DistanceFromBottomReferenceMeters { get; set; }

        [JsonProperty("IsLongNozzle")]
        public bool IsLongNozzle { get; set; }

        [JsonProperty("TopReferenceType")]
        public NozzleTopReferenceType TopReferenceType { get; set; }

        [JsonProperty("DistanceFromTopReferenceMeters")]
        public double DistanceFromTopReferenceMeters { get; set; }

        [JsonProperty("DistanceFromReference")]
        public double DistanceFromReference { get; set; }

        [JsonProperty("ReferenceType")]
        public NozzleReferenceType ReferenceType { get; set; }

        [JsonProperty("IsReferenceToLeft")]
        public bool IsReferenceToLeft { get; set; }

        public NozzleSettings() { }

        public NozzleSettings DeepClone()
        {
            return new NozzleSettings
            {
                ID = this.ID,
                PIDCenterAxis = this.PIDCenterAxis?.ToArray(),
                PIDNozzleAxis = this.PIDNozzleAxis?.ToArray(),
                PIDPositionPlane = this.PIDPositionPlane?.ToArray(),
                PIDExternalPoint = this.PIDExternalPoint?.ToArray(),
                PIDInternalPoint = this.PIDInternalPoint?.ToArray(),
                PIDInsidePoint = this.PIDInsidePoint?.ToArray(),
                PIDMidPoint = this.PIDMidPoint?.ToArray(),
                PIDTopPoint = this.PIDTopPoint?.ToArray(),
                PIDNozzleRightRefPlane = this.PIDNozzleRightRefPlane?.ToArray(),
                PIDCutPlane = this.PIDCutPlane?.ToArray(),
                PIDSketch = this.PIDSketch?.ToArray(),
                PIDComponent = this.PIDComponent?.ToArray(),
                PIDPositionPlaneMate = this.PIDPositionPlaneMate?.ToArray(),
                PIDTopPlaneMate = this.PIDTopPlaneMate?.ToArray(),
                PIDNozzleAssemblyComp = this.PIDNozzleAssemblyComp?.ToArray(),
                PIDCutOutPlane = this.PIDCutOutPlane?.ToArray(),
                PIDCutExtrude = this.PIDCutExtrude?.ToArray(),
                PIDCenterlineWallIntersection = this.PIDCenterlineWallIntersection?.ToArray(),
                NozzleAssemblyComponents = this.NozzleAssemblyComponents,
                RotationAngleDegrees = this.RotationAngleDegrees,
                OffsetMeters = this.OffsetMeters,
                OffsetPosition = this.OffsetPosition,
                BottomReferencePoint = this.BottomReferencePoint,
                DistanceFromBottomReferenceMeters = this.DistanceFromBottomReferenceMeters,
                IsLongNozzle = this.IsLongNozzle,
                TopReferenceType = this.TopReferenceType,
                DistanceFromTopReferenceMeters = this.DistanceFromTopReferenceMeters,
                DistanceFromReference = this.DistanceFromReference,
                ReferenceType = this.ReferenceType,
                IsReferenceToLeft = this.IsReferenceToLeft
            };
        }
    }
}

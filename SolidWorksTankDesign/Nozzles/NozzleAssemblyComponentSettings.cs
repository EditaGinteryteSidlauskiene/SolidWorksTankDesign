using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    public class NozzleAssemblyComponentSettings
    {
        [JsonProperty("PIDComponent")]
        public byte[] PIDComponent { get; set; }

        [JsonProperty("PIDMatingPlane")]
        public byte[] PIDMatingPlane { get; set; }

        [JsonProperty("PIDFreePlane")]
        public byte[] PIDFreePlane { get; set; }

        [JsonProperty("ComponentType")]
        public NozzleComponentType ComponentType { get; set; }

        [JsonProperty("HasAdjustableLength")]
        public bool HasAdjustableLength { get; set; }

        /// <summary>
        /// The gap between this component's mating plane and the free plane of the component
        /// above it (e.g. a 2 mm welding gap). Included in total nozzle length calculations.
        /// Read from the SolidWorks custom property "WeldingGapMeters" when the assembly is added,
        /// or set explicitly when components are added programmatically.
        /// </summary>
        [JsonProperty("WeldingGapMeters")]
        public double WeldingGapMeters { get; set; }

        [JsonProperty("Diameter")]
        public double Diameter { get; set; }

        [JsonProperty("IsShellIntersecting")]
        public bool IsShellIntersecting { get; set; }
    }
}

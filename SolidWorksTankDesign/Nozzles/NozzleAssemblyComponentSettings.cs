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

        [JsonProperty("Diameter")]
        public double Diameter { get; set; }

        [JsonProperty("IsShellIntersecting")]
        public bool IsShellIntersecting { get; set; }
    }
}

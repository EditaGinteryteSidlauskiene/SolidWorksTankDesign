using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    internal class AssemblyOfCylindricalShells
    {
        [JsonProperty("CylindricalShells")]
        public List<CylindricalShell> CylindricalShells = new List<CylindricalShell>();

        public Feature GetCenterAxis() => FeatureManager.GetFeatureByName(SolidWorksDocumentProvider.ActiveDoc(), "Center axis");

        public AssemblyOfCylindricalShells() { }

    }
}

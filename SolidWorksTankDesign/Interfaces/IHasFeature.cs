using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// Represents an entity that has an associated SolidWorks feature.
    /// </summary>
    internal interface IHasFeature
    {
        /// <summary>
        /// Gets the associated SolidWorks feature.
        /// </summary>
        Feature Feature { get; }

        void Set(Feature feature);
    }
}

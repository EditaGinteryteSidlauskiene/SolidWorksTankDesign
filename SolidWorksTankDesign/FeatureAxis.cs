using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// Represents a SolidWorks reference axis feature, providing access to its associated Feature, 
    /// RefAxis, and RefAxisFeatureData objects.
    /// </summary>
    internal class FeatureAxis : IHasFeature
    {
        /// <summary>
        /// Gets the underlying SolidWorks feature object.
        /// </summary>
        public Feature Feature { get; set; }

        /// <summary>
        /// Gets the SolidWorks reference axis object.
        /// </summary>
        internal RefAxis RefAxis { get; private set; }

        /// <summary>
        /// Gets the reference axis feature data, providing access to specific axis properties.
        /// </summary>
        internal RefAxisFeatureData RefAxisFeatureData { get; private set; }

        /// <summary>
        /// Updates the internal Feature, RefAxis, and RefAxisFeatureData based on a provided Feature.
        /// </summary>
        /// <param name="feature">The new SolidWorks Feature to associate with this FeatureAxis.</param>
        public void Set(Feature feature)
        {
            Feature = feature;
            RefAxis = Feature.GetSpecificFeature2();
            RefAxisFeatureData = Feature.GetDefinition();
        }
    }
}

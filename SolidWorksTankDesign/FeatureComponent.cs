using SolidWorks.Interop.sldworks;

namespace SolidWorksTankDesign
{
    internal class FeatureComponent : IHasFeature
    {
        /// <summary>
        /// Gets the underlying SolidWorks feature object.
        /// </summary>
        public Feature Feature { get; private set; }

        /// <summary>
        /// Gets the SolidWorks component object.
        /// </summary>
        internal Component2 Component { get; /*private*/ set; }

        /// <summary>
        /// Updates the internal Feature and Component based on a provided Feature.
        /// </summary>
        /// <param name="feature">The new SolidWorks Feature to associate with this FeatureComponent.</param>
        public void Set(Feature feature)
        {
            Feature = feature;
            Component = Feature.GetSpecificFeature2();
        }
    }
}

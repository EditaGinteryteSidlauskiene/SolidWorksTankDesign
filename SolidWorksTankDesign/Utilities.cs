using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SolidWorksTankDesign
{
    internal static class Utilities
    {
        /// <summary>
        /// Dictionary of Feature Type Names. It can be used to get Feature TypeName as a string
        /// </summary>
        /// 
        static readonly Dictionary<FeatureType, string> FeatureTypeName = new Dictionary<FeatureType, string>()
        {
            {FeatureType.RefPlane, "RefPlane" },
            {FeatureType.RefAxis, "RefAxis" },
            {FeatureType.RefPoint, "RefPoint" },
            {FeatureType.Component, "Reference" },
            {FeatureType.Sketch, "ProfileFeature" },
            {FeatureType.Attribute, "Attribute" }
        };

        /// <summary>
        /// Extracts top-level components from a SolidWorks assembly document.
        /// </summary>
        /// <param name="Assembly">The SolidWorks assembly document (ModelDoc2) to process.</param>
        /// <returns>A List where with Component2 objects.</returns>
        public static List<Component2> GetTopLevelComponents(ModelDoc2 Document) => GetAllComponents(Document).Select(fs => fs.Component).ToList();

        /// <summary>
        /// Retrieves all component features from a ModelDoc2 document.
        /// </summary>
        /// <param name="Document">The ModelDoc2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns> 
        public static List<FeatureComponent> GetAllComponents(ModelDoc2 Document) => GetAllFeaturesOfType<FeatureComponent>(Document.IFirstFeature(), FeatureType.Component);

        /// <summary>
        /// Retrieves all component features from a Component2 document.
        /// </summary>
        /// <param name="Document">The Component2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns>
        /// 
        public static List<FeatureComponent> GetAllComponents(Component2 Document) => GetAllFeaturesOfType<FeatureComponent>(Document.FirstFeature(), FeatureType.Component);

        /// <summary>
        /// Retrieves all reference axis features from a ModelDoc2 document.
        /// </summary>
        /// <param name="Document">The ModelDoc2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns> 
        public static List<FeatureAxis> GetAllReferenceAxisFeatures(ModelDoc2 Document) => GetAllFeaturesOfType<FeatureAxis>(Document.IFirstFeature(), FeatureType.RefAxis);

        /// <summary>
        /// Retrieves all features of a specified type from a given starting feature.
        /// </summary>
        /// <param name="Feature">The feature to start the search from.</param>
        /// <param name="FeatureType">The type of feature to retrieve.</param>
        /// <returns>A list of all features matching the specified type.</returns>
        private static List<T> GetAllFeaturesOfType<T>(Feature Feature, FeatureType FeatureType) where T : IHasFeature, new()
        {
            List<T> returnCollection = new List<T>();

            string featureTypeName = FeatureTypeName[FeatureType];

            while (Feature != null)
            {
                if (featureTypeName == Feature.GetTypeName2())
                {
                    T IHasFeatureObject = new T();
                    IHasFeatureObject.Set(Feature);

                    returnCollection.Add(IHasFeatureObject);
                }

                Feature = (Feature)Feature.GetNextFeature();
            }

            return returnCollection; // Return the collected features
        }

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="ModelDocument">The model document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(ModelDoc2 ModelDocument, FeatureType DesiredFeatureType, int Count) => GetNthFeatureOfType(ModelDocument.IFirstFeature(), DesiredFeatureType, Count);

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="Component">The component document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(Component2 Component, FeatureType DesiredFeatureType, int Count) => GetNthFeatureOfType(Component.FirstFeature(), DesiredFeatureType, Count);

        private static Feature GetNthFeatureOfType(Feature FirstFeature, FeatureType DesiredFeatureType, int Count)
        {
            // Error Handling: Ensure the provided count is valid
            if (Count <= 0)
            {
                return null;
            }

            // Getting the desired feature type's name for comparison
            string featureTypeName = FeatureTypeName[DesiredFeatureType];

            // Initialize variables 
            Feature loopFeature = FirstFeature;
            int featureCounter = 0;

            // Iterate through features 
            while (loopFeature != null)
            {
                if (featureTypeName == loopFeature.GetTypeName2())
                {
                    featureCounter++;

                    if (featureCounter == Count)
                        return loopFeature;
                }

                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }
    }
}
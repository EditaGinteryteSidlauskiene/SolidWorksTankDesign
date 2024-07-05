using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace SolidWorksTankDesign
{
    internal static class FeatureManager
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
        public static List<Component2> GetTopLevelComponents(ModelDoc2 document) => GetAllComponents(document).Select(fs => fs.Component).ToList();

        /// <summary>
        /// Retrieves all component features from a ModelDoc2 document.
        /// </summary>
        /// <param name="document">The ModelDoc2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns> 
        public static List<FeatureComponent> GetAllComponents(ModelDoc2 document) => GetAllFeaturesOfType<FeatureComponent>(document.IFirstFeature(), FeatureType.Component);

        /// <summary>
        /// Retrieves all component features from a Component2 document.
        /// </summary>
        /// <param name="document">The Component2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns>
        /// 
        public static List<FeatureComponent> GetAllComponents(Component2 document) => GetAllFeaturesOfType<FeatureComponent>(document.FirstFeature(), FeatureType.Component);

        /// <summary>
        /// Searches for a child component with the specified name within the given parent component.
        /// </summary>
        /// <param name="parentComponent">The component whose children will be searched.</param>
        /// <param name="componentToFindName">The name of the child component to find.</param>
        /// <returns>The child component with the matching name, or null if not found.</returns>
        /// <remarks>
        /// This method iterates through all child components of the parent component. If a component with the 
        /// specified name is found, it is immediately returned. If not found, null is returned.
        /// </remarks>
        public static Component2 GetChildComponentByName(Component2 parentComponent, string componentToFindName)
        {
            // Get a list of all feature components under the parent component
            List<FeatureComponent> allComponents = FeatureManager.GetAllComponents(parentComponent);

            // Iterate through each feature component
            foreach (FeatureComponent currentComponent in allComponents)
            {
                // Check if the component's name matches the desired name
                if (currentComponent.Feature.Name == componentToFindName) return currentComponent.Component;
            }
            // Component not found, return null
            return null;
        }

        /// <summary>
        /// Retrieves all reference axis features from a ModelDoc2 document.
        /// </summary>
        /// <param name="document">The ModelDoc2 document to search.</param>
        /// <returns>A list of all reference axis features found in the document.</returns> 
        public static List<FeatureAxis> GetAllReferenceAxisFeatures(ModelDoc2 document) => GetAllFeaturesOfType<FeatureAxis>(document.IFirstFeature(), FeatureType.RefAxis);

        /// <summary>
        /// Retrieves all features of a specified type from a given starting feature.
        /// </summary>
        /// <param name="feature">The feature to start the search from.</param>
        /// <param name="featureType">The type of feature to retrieve.</param>
        /// <returns>A list of all features matching the specified type.</returns>
        private static List<T> GetAllFeaturesOfType<T>(Feature feature, FeatureType featureType) where T : IHasFeature, new()
        {
            List<T> returnCollection = new List<T>();

            string featureTypeName = FeatureTypeName[featureType];

            while (feature != null)
            {
                if (featureTypeName == feature.GetTypeName2())
                {
                    T IHasFeatureObject = new T();
                    IHasFeatureObject.Set(feature);

                    returnCollection.Add(IHasFeatureObject);
                }

                feature = (Feature)feature.GetNextFeature();
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
        /// <param name="component">The component document containing the features.</param>
        /// <param name="desiredFeatureType">The desired feature type to search for.</param>
        /// <param name="count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(Component2 component, FeatureType desiredFeatureType, int count) => GetNthFeatureOfType(component.FirstFeature(), desiredFeatureType, count);

        private static Feature GetNthFeatureOfType(Feature firstFeature, FeatureType desiredFeatureType, int Count)
        {
            // Error Handling: Ensure the provided count is valid
            if (Count <= 0)
            {
                return null;
            }

            // Getting the desired feature type's name for comparison
            string featureTypeName = FeatureTypeName[desiredFeatureType];

            // Initialize variables 
            Feature loopFeature = firstFeature;
            int featureCounter = 0;

            // Iterate through features 
            while (loopFeature != null)
            {
                Debug.WriteLine(loopFeature.Name + " " + loopFeature.GetTypeName2());

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

        /// <summary>
        /// Gets feature of the document by name by iterating all features until the requested one is reached.
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// 
        public static Feature GetFeatureByName(ModelDoc2 modelDocument, string name)
        {
            //Starting from the first feature
            Feature loopFeature = modelDocument.IFirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.Name == name)
                {
                    return loopFeature;
                }

                //Get next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }

        /// <summary>
        ///  Gets feature of the component by name by iterating all features until the requested one is reached.
        /// </summary>
        /// <param name="component"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Feature GetFeatureByName(Component2 component, string name)
        {
            //Starting from the first feature
            Feature loopFeature = component.FirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.Name == name)
                {
                    return loopFeature;
                }

                //Get next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }

    }
}

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml.Linq;

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
        /// Dictionary of Entity Types. These type names are used in SelectByID2 method. The key in 
        /// this dictionary is feature's type name.
        /// </summary>
        public static Dictionary<string, string> EntityType = new Dictionary<string, string>()
        {
            {"RefPlane", "PLANE" },
            {"RefAxis", "AXIS" },
            {"MateCoincident", "MATE" },
            {"MateDistanceDim", "MATE"},
            {"Reference", "COMPONENT" }
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
        /// This method programmatically selects a specified component within a SolidWorks model document.
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="component"></param>
        public static void SelectComponent(ModelDoc2 modelDocument, Component2 component)
        {
            ModelDocExtension swModelDocExt = modelDocument.Extension;

            // Select assembly's feature
            swModelDocExt.SelectByID2(
                Name: component.Name + "@" + modelDocument.GetTitle(),
                Type: "COMPONENT",
                X: 0,
                Y: 0,
                Z: 0,
                Append: true,
                Mark: 1,
                Callout: null,
                SelectOption: (int)swSelectOption_e.swSelectOptionDefault);
        }

        /// <summary>
        /// This method programmatically selects a specified feature within a SolidWorks model document.
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="feature"></param>
        public static void SelectFeature(ModelDoc2 modelDocument, Feature feature)
        {
            ModelDocExtension swModelDocExt = modelDocument.Extension;

            //Setting feature type name
            string featureTypeName = EntityType[feature.GetTypeName2()];

            // Select assembly's feature
            swModelDocExt.SelectByID2(
                Name: feature.Name,
                Type: featureTypeName,
                X: 0,
                Y: 0,
                Z: 0,
                Append: true,
                Mark: 1,
                Callout: null,
                SelectOption: (int)swSelectOption_e.swSelectOptionDefault);
        }

        /// <summary>
        /// This method programmatically selects a specific feature of a component within a SolidWorks model document.
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="component"></param>
        /// <param name="componentFeature"></param>
        public static void SelectFeature(ModelDoc2 modelDocument, Component2 component, Feature componentFeature)
        {
            ModelDocExtension swModelDocExt = modelDocument.Extension;

            //Getting full entity name, which will be used later to select the entity. Example of full entity name
            //FullEntityName = "Top@" + strComponentName + "@" + AssemblyName;
            string fullComponentFeatureName =
                componentFeature.Name
                + "@"
                + component.Name
                + "@"
                + modelDocument.GetTitle();

            //Setting entity type name
            string componentFeatureTypeName = EntityType[componentFeature.GetTypeName2()];

            // Select component's feature
            swModelDocExt.SelectByID2(
                Name: fullComponentFeatureName,
                Type: componentFeatureTypeName,
                X: 0,
                Y: 0,
                Z: 0,
                Append: true,
                Mark: 1,
                Callout: null,
                SelectOption: (int)swSelectOption_e.swSelectOptionDefault);
        }

        /// <summary>
        /// This method retrieves the persistent reference ID of a SolidWorks entity (component, feature, etc.) within a given model document.
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="entity"></param>
        /// <returns></returns>
        public static byte[] GetPersistentReferenceId(ModelDocExtension modelDocExtension, object entity)
        {
            // Safety Check: Ensure the input entity is not null
            if (entity == null)
            {
                throw new ArgumentNullException(nameof(entity), "The SolidWorks entity cannot be null.");
            }

            // Retrieve the persistent reference ID of the entity
            byte[] persistentReferenceId = modelDocExtension.GetPersistReference3(entity);

            // Safety Check: Ensure a valid persistent ID was obtained
            if (persistentReferenceId == null || persistentReferenceId.Length == 0)
            {
                throw new InvalidOperationException("Failed to retrieve a valid persistent reference ID.");
            }


            return persistentReferenceId;
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
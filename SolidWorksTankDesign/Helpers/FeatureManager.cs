﻿using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

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
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="ModelDocument">The model document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(ModelDoc2 modelDocument, FeatureType desiredFeatureType, int count) => 
            GetNthFeatureOfType(modelDocument.IFirstFeature(), desiredFeatureType, count);

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="Component">The component document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(Component2 component, FeatureType desiredFeatureType, int count) => GetNthFeatureOfType(component.FirstFeature(), desiredFeatureType, count);

        /// <summary>
        /// Returns a major plane of ModelDoc2 document that is requested by providing its type from MajorPlane enum.
        /// </summary>
        /// <param name="PlaneType"></param>
        /// <param name="Document"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(ModelDoc2 document, MajorPlane planeType) =>
            GetNthFeatureOfType(document, FeatureType.RefPlane, (int)planeType);

        /// <summary>
        /// Returns a major plane of Component2 document that is requested by providing its type from MajorPlane enum.
        /// </summary>
        /// <param name="PlaneType"></param>
        /// <param name="Component"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(Component2 component, MajorPlane planeType) => GetNthFeatureOfType(component, FeatureType.RefPlane, (int)planeType);

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
        /// Gets feature of the document by name by iterating all features until the requested one is reached.
        /// </summary>
        /// <param name="Component"></param>
        /// <param name="Name"></param>
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

        /// <summary>
        /// Creates a new reference plane in a SolidWorks model, a specified distance away from an existing plane.
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="ExistingPlane"></param>
        /// <param name="Distance"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public static Feature CreateReferencePlaneWithDistance(
            ModelDoc2 modelDocument,
            Feature existingPlane,
            double distance,
            string name)
        {
            //Selects the existing plane 
            bool status = existingPlane.Select2(false, 0);

            //Creates a new reference plane
            Feature ReferencePlane = (Feature)modelDocument.FeatureManager.InsertRefPlane(8, distance, 0, 0, 0, 0);

            //Rename just created reference plane
            RenameFeature(ReferencePlane, name);

            return ReferencePlane;
        }

        /// <summary>
        /// Renames feature
        /// </summary>
        /// <param name="swFeature"></param>
        /// <param name="Name"></param>
        public static void RenameFeature(Feature swFeature, string name) => swFeature.Name = name;

        /// <summary>
        /// Suppress feature
        /// </summary>
        public static void Suppress(Feature featureToSuppress)
        {
            ChangeSuppression(featureToSuppress, 0);
        }

        /// <summary>
        /// Unsuppress feature
        /// </summary>
        public static void Unsuppress(Feature featureToUnsuppress)
        {
            ChangeSuppression(featureToUnsuppress, 1);
        }

        /// <summary>
        /// Change the state of suppression. Suppression State = 0 to suppress, 1 to unsuppress.
        /// </summary>
        /// <param name="feature"></param>
        /// <param name="suppressionState"></param>
        private static void ChangeSuppression(Feature feature, int suppressionState)
        {
            feature.SetSuppression2(suppressionState, 2, "");
        }

        /// <summary>
        /// Changes reference entity (NewReferencePlane) for reference plane feature (ReferencePlane).
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="newReferencePlane"></param>
        /// <param name="referencedPlane"></param>
        public static bool ChangeReferenceOfReferencePlane(
            ModelDoc2 modelDocument,
            Feature newReferencePlane,
            Feature referencedPlane)
        {
            //Get access to reference plane properties
            RefPlaneFeatureData referencePlaneFeatureData = referencedPlane.GetDefinition();

            //Set new reference plane
            referencePlaneFeatureData.Reference[0] = newReferencePlane;

            //Modify changes
            return referencedPlane.ModifyDefinition(referencePlaneFeatureData, modelDocument, null);
        }

        /// <summary>
        /// Changes the distance of reference plane from the starting plane
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="referencePlane"></param>
        /// <param name="distance"></param>
        public static bool ChangeDistanceOfReferencePlane(
            Feature referencePlane,
            double distance)
        {
            try
            {
                //Get access to reference plane properties
                RefPlaneFeatureData referencePlaneFeatureData = referencePlane.GetDefinition();

                //Set new distance
                referencePlaneFeatureData.Distance = distance;

                //Modify changes
                return referencePlane.ModifyDefinition(referencePlaneFeatureData, SolidWorksDocumentProvider.ActiveDoc(), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Distance of reference plane could not be changed: {ex.Message}");
                return false;
            }
        }
    }
}

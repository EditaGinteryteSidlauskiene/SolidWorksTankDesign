using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal static class SWFeatureManager
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
        /// Finds and returns a specific feature within a sequence of features in a SolidWorks model or assembly.
        /// </summary>
        /// <param name="firstFeature"></param>
        /// <param name="desiredFeatureType"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private static Feature GetNthFeatureOfType(Feature firstFeature, FeatureType desiredFeatureType, int count)
        {
            // Error Handling: Ensure the provided count is valid
            if (count <= 0)
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

                    if (featureCounter == count)
                        return loopFeature;
                }

                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
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
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="modelDocument">The model document containing the features.</param>
        /// <param name="desiredFeatureType">The desired feature type to search for.</param>
        /// <param name="count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(ModelDoc2 modelDocument, FeatureType desiredFeatureType, int count) => 
            GetNthFeatureOfType(modelDocument.IFirstFeature(), desiredFeatureType, count);

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="component">The component document containing the features.</param>
        /// <param name="desiredFeatureType">The desired feature type to search for.</param>
        /// <param name="count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(Component2 component, FeatureType desiredFeatureType, int count) => GetNthFeatureOfType(component.FirstFeature(), desiredFeatureType, count);

        /// <summary>
        /// Returns a major plane of ModelDoc2 document that is requested by providing its type from MajorPlane enum.
        /// </summary>
        /// <param name="planeType"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(ModelDoc2 document, MajorPlane planeType) =>
            GetNthFeatureOfType(document, FeatureType.RefPlane, (int)planeType);

        /// <summary>
        /// Returns a major plane of Component2 document that is requested by providing its type from MajorPlane enum.
        /// </summary>
        /// <param name="planeType"></param>
        /// <param name="component"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(Component2 component, MajorPlane planeType) => GetNthFeatureOfType(component, FeatureType.RefPlane, (int)planeType);

        /// <summary>
        /// Gets feature of the document by name by iterating all features until the requested one is reached.
        /// </summary>
        /// <param name="assemblyModelDoc"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// 
        public static Feature GetFeatureByName(ModelDoc2 assemblyModelDoc, string name)
        {
            //Starting from the first feature
            Feature loopFeature = assemblyModelDoc.IFirstFeature();

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

        /// <summary>
        /// Creates a new reference plane in a SolidWorks model, a specified distance away from an existing plane.
        /// </summary>
        /// <param name="existingPlane"></param>
        /// <param name="distance"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Feature CreateReferencePlaneWithDistance(
            Feature existingPlane,
            double distance,
            string name,
            bool flip)
        {
            //Selects the existing plane 
            existingPlane.Select2(false, 0);

            //Creates a new reference plane
            Feature referencePlane = (Feature)SolidWorksDocumentProvider.GetActiveDoc().FeatureManager.InsertRefPlane(8, distance, 0, 0, 0, 0);

            if(flip)
            {
                //Get access to reference plane properties
                RefPlaneFeatureData referencePlaneFeatureData = (RefPlaneFeatureData)referencePlane.GetDefinition();

                //Set new reference plane
                referencePlaneFeatureData.ReversedReferenceDirection[0] = true;

                //Modify changes
                referencePlane.ModifyDefinition(
                    referencePlaneFeatureData, 
                    SolidWorksDocumentProvider.GetActiveDoc(), 
                    null);
            }

            //Rename just created reference plane
            RenameFeature(referencePlane, name);

            return referencePlane;
        }

        /// <summary>
        /// Renames feature
        /// </summary>
        /// <param name="swFeature"></param>
        /// <param name="name"></param>
        public static void RenameFeature(Feature swFeature, string name) => swFeature.Name = name;

        /// <summary>
        /// Suppress feature
        /// </summary>
        public static void Suppress(Feature featureToSuppress)
        {
            ChangeSuppression(featureToSuppress, 0);
        }

        /// <summary>
        /// Suppress component
        /// </summary>
        public static void Suppress(Component2 componentToSuppress)
        {
            componentToSuppress.SetSuppression2((int)swComponentSuppressionState_e.swComponentSuppressed);
        }

        /// <summary>
        /// Unsuppress feature
        /// </summary>
        public static void Unsuppress(Feature featureToUnsuppress)
        {
            ChangeSuppression(featureToUnsuppress, 1);
        }

        /// <summary>
        /// Unsuppress component
        /// </summary>
        public static void Unsuppress(Component2 componentToUnsuppress)
        {
            componentToUnsuppress.SetSuppression2((int)swComponentSuppressionState_e.swComponentFullyResolved);
        }

        /// <summary>
        /// Changes reference entity (NewReferencePlane) for reference plane feature (ReferencePlane).
        /// </summary>
        /// <param name="modelDocument"></param>
        /// <param name="newReferencePlane"></param>
        /// <param name="referencedPlane"></param>
        public static bool ChangeReferenceOfReferencePlane(
            Feature newReferencePlane,
            Feature referencedPlane)
        {
            // Currently active model doc
            ModelDoc2 activeDoc = SolidWorksDocumentProvider.GetActiveDoc();

            //Get access to reference plane properties
            RefPlaneFeatureData referencePlaneFeatureData = referencedPlane.GetDefinition();

            //Set new reference plane
            referencePlaneFeatureData.Reference[0] = newReferencePlane;

            //Modify changes
            bool success =  referencedPlane.ModifyDefinition(referencePlaneFeatureData, activeDoc, null);

            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.ActivateDocument();
            ComponentManager.RefreshDishedEnds();

            int errors = 0;
            int warnings = 0;
            SolidWorksDocumentProvider.GetActiveDoc().Save3(
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                            ref errors,
                            ref warnings);

            return success;
        }

        public static double GetDistanceOfReferencePlane(Feature referencePlane)
        {
            try
            {
                //Get access to reference plane properties
                RefPlaneFeatureData referencePlaneFeatureData = referencePlane.GetDefinition();

                //Set new distance
                return referencePlaneFeatureData.Distance;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }
        }

        /// <summary>
        /// Changes the distance of reference plane from the starting plane
        /// </summary>
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
                return referencePlane.ModifyDefinition(referencePlaneFeatureData, SolidWorksDocumentProvider.GetActiveDoc(), null);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Distance of reference plane could not be changed: {ex.Message}");
                return false;
            }
        }

        public static Feature GetFeature(byte[] PIDFeature, byte[] PIDComponent2, byte[] PIDComponent1)
        {
            // Activate component's doc
            Component2 component1FromPID = (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                       PIDComponent1,
                       out int error);

            ModelDoc2 component1Doc = component1FromPID.GetModelDoc2();

            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(component1Doc.GetTitle(), true, 0, 0);

            // Activate child component's doc
            Component2 component2FromPID = (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                       PIDComponent2,
                       out error);

            // Close parent doc
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(component1Doc.GetTitle());

           ModelDoc2 component2Doc = component2FromPID.GetModelDoc2();

            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(component2Doc.GetTitle(), true, 0, 0);

            // Get feature from PID
            return (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                       PIDFeature,
                       out error);
        }

        public static double GetCustomPropertyValue(string documentPath, string customPropertyName)
        {
            // Get cylindrical shell's volume per 1 meter
            int error = 0;
            int warning = 0;

            ModelDoc2 document = SolidWorksDocumentProvider._solidWorksApplication.OpenDoc6(
                documentPath,
                (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref error,
                ref warning);

            CustomPropertyManager customPropertyManager = document.Extension.get_CustomPropertyManager("");

            customPropertyManager.Get6(
                customPropertyName,
                false,
                out string volumePerMeter,
                out _, out _, out _);

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(document.GetTitle());

            return double.Parse(volumePerMeter.Replace(',', '.'));
        }

        /// <summary>
        /// Reads a string custom property directly from an already-open component's model document.
        /// </summary>
        public static string GetCustomPropertyFromComponent(Component2 component, string propertyName)
        {
            ModelDoc2 compDoc = component.GetModelDoc2();
            if (compDoc == null) return null;

            CustomPropertyManager manager = compDoc.Extension.get_CustomPropertyManager("");
            manager.Get6(propertyName, false, out string val, out _, out _, out _);

            return val;
        }

        /// <summary>
        /// CURRENTLY ACTIVE SHELL DOC MUST BE ACTIVATED!!!
        /// </summary>
        /// <param name="dishedEnd"></param>
        /// <returns></returns>
        public static Feature GetDishedEndPositionPlane(DishedEnd dishedEnd)
        {
            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();
            ModelDoc2 currentlyActiveShellDoc = SolidWorksDocumentProvider.GetActiveDoc();
            //Left end plane@Assembly of Dished ends-1@Shell
            // Get position plane's name from assembly of dished ends perspective
            string positionPlaneName = GetPositionPlaneName(dishedEnd);

            SelectionMgr selectionMgr = currentlyActiveShellDoc.SelectionManager;

            // Select position plane using this name
            currentlyActiveShellDoc.Extension.SelectByID2(positionPlaneName, "PLANE", 0, 0, 0, false, 0, null, 0);

            // Get selected object
            return selectionMgr.GetSelectedObject6(1, -1);

            string GetPositionPlaneName(DishedEnd dishedEndObject)
            {
                // 1. Get assembly of dished ends document
                Component2 assemblyOfDishedEndsComp = SolidWorksDocumentProvider._tankSiteAssembly.GetDishedEndsAssemblyComponent();
                ModelDoc2 assemblyOfDishedEndsDoc = assemblyOfDishedEndsComp.GetModelDoc2();

                // 2. Activate assembly of dished ends doc to be able to get position plane
                SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfDishedEndsDoc.GetTitle(), true, 0, 0);

                // Get shell component name
                string shellCompFullName = SolidWorksDocumentProvider._tankSiteAssembly.GetShellAssemblyComponent().Name2;
                string[] shellNameParts = shellCompFullName.Split('/');
                string shellName = shellNameParts[shellNameParts.Length - 1].Split('-')[0];

                // Get dished end's component name
                string dishedEndsCompFullName = assemblyOfDishedEndsComp.Name2;
                string[] dishedEndNameParts = dishedEndsCompFullName.Split('/');
                string dishedEndsCompName = dishedEndNameParts[dishedEndNameParts.Length - 1];

                // 3. Get position plane
                Feature positionPlaneOfDishedEnd = dishedEndObject.GetPositionPlane();

                // Close assembly of dished ends doc
                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfDishedEndsDoc.GetTitle());

                // 4. Get position plane's name from assembly of dished ends doc perspective
                return $"{positionPlaneOfDishedEnd.Name}@{dishedEndsCompName}@{shellName}";
            }
        }

        public static void UpdateMatePID(ModelDoc2 document, Compartment compartment, string mateName)
        {
            Feature mate = MateManager.GetMateByName(document, mateName);
            compartment._compartmentSettings.PIDLeftEndMate = document.Extension.GetPersistReference3(mate);
        }

        /// <summary>
        /// Checks if a plane's normal vector is pointing in the positive Z direction (upward).
        /// Returns true if the plane normal points upward (positive Z), false if it points downward (negative Z).
        /// </summary>
        /// <param name="plane">The plane feature to check</param>
        /// <returns>True if plane normal points upward, false if downward</returns>
        public static bool IsPlaneNormalPointingUp(Feature plane)
        {
            try
            {
                if (plane == null) return true; // Default to true if plane is null

                RefPlane refPlane = (RefPlane)plane.GetSpecificFeature2();
                if (refPlane == null) return true;

                MathTransform planeTransform = refPlane.Transform;
                if (planeTransform == null) return true;

                // Get the Z-axis (normal) of the plane's coordinate system
                double[] zVector = new double[] { 0, 0, 1 };
                SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
                SolidWorks.Interop.sldworks.MathUtility mathUtil = (SolidWorks.Interop.sldworks.MathUtility)solidWorksApp.GetMathUtility();

                SolidWorks.Interop.sldworks.MathVector normalVectorObj = (SolidWorks.Interop.sldworks.MathVector)mathUtil.CreateVector(zVector);
                normalVectorObj = (SolidWorks.Interop.sldworks.MathVector)normalVectorObj.MultiplyTransform(planeTransform);

                double[] planeNormal = (double[])normalVectorObj.ArrayData;

                // Check if Z component is positive (pointing upward)
                // If Z > 0, plane is pointing up. If Z < 0, plane is pointing down.
                return planeNormal[2] > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking plane normal direction: {ex.Message}");
                return true; // Default to true on error
            }
        }
    }
}

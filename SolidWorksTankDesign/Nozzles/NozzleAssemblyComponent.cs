using SolidWorks.Interop.sldworks;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// Represents a single component (flange, pipe, elbow, etc.) within a nozzle assembly.
    /// Supports flexible configuration for different component types with varying connection points.
    /// </summary>
    public class NozzleAssemblyComponent
    {
        [JsonProperty("Settings")]
        public NozzleAssemblyComponentSettings Settings { get; set; } = new NozzleAssemblyComponentSettings();

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     Settings.PIDComponent,
                     out int error);

        public Feature GetMatingPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     Settings.PIDMatingPlane,
                     out int error);

        public Feature GetFreePlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     Settings.PIDFreePlane,
                     out int error);

        /// <summary>
        /// Returns the current D1 value (metres) from the component's "Path Sketch".
        /// Activates and closes the component doc internally.
        /// </summary>
        public double GetCurrentLength()
        {
            Component2 nozzleComp = GetComponent();
            if (nozzleComp == null) return 0;

            ModelDoc2 compDoc = nozzleComp.GetModelDoc2();
            if (compDoc == null) return 0;

            int errors = 0;
            compDoc = (ModelDoc2)SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(
                compDoc.GetPathName(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref errors);
            if (compDoc == null) return 0;

            Feature pathSketch = SWFeatureManager.GetFeatureByName(compDoc, "Path Sketch");
            if (pathSketch == null) return 0;

            Dimension d1 = pathSketch.Parameter("D1");
            if (d1 == null) return 0;

            double value = ((double[])d1.GetSystemValue3(
                (int)swInConfigurationOpts_e.swThisConfiguration, null))[0];

            // Close the component doc so the nozzle doc remains the active context.
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(compDoc.GetTitle());

            return value;
        }

        public void ChangeLength(double newLength)
        {
            Component2 nozzleComp = GetComponent();
            if (nozzleComp == null)
            {
                MessageBox.Show("Component not found.");
                return;
            }

            ModelDoc2 compDoc = nozzleComp.GetModelDoc2();
            if (compDoc == null)
            {
                MessageBox.Show("Component document not found.");
                return;
            }

            // Activate the component's document
            int errors = 0;
            int warnings = 0;
            compDoc = (ModelDoc2)SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(
                compDoc.GetPathName(), false, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, ref errors);

            if (compDoc == null)
            {
                MessageBox.Show("Failed to activate component document.");
                return;
            }

            // Get D1 from "Path Sketch"
            Feature pathSketch = SWFeatureManager.GetFeatureByName(compDoc, "Path Sketch");
            if (pathSketch == null)
            {
                MessageBox.Show("\"Path Sketch\" feature not found in component.");
                return;
            }

            Dimension d1 = pathSketch.Parameter("D1");
            if (d1 == null)
            {
                MessageBox.Show("D1 dimension not found in \"Path Sketch\".");
                return;
            }

            d1.SetSystemValue3(
                newLength,
                (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration,
                null);

            compDoc.EditRebuild3();

            // Close the component doc so the nozzle doc remains the active context.
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(compDoc.GetTitle());
        }

        /// <summary>
        /// Determines whether this component intersects the shell by checking if the external point
        /// (where the nozzle meets the outer shell wall) lies between this component's mating plane
        /// and free plane. Works regardless of how the planes are oriented per component type.
        /// </summary>
        /// <param name="externalPointCoords">XYZ coordinates of the external point in meters</param>
        /// <returns>True if the external point is between the mating and free planes</returns>
        public bool IsShellIntersecting(double[] externalPointCoords)
        {
            try
            {
                Feature matingPlaneFeature = GetMatingPlane();
                Feature freePlaneFeature   = GetFreePlane();

                if (matingPlaneFeature == null || freePlaneFeature == null)
                    return false;

                RefPlane matingPlane = (RefPlane)matingPlaneFeature.GetSpecificFeature2();
                RefPlane freePlane   = (RefPlane)freePlaneFeature.GetSpecificFeature2();

                if (matingPlane == null || freePlane == null)
                    return false;

                Component2 comp = GetComponent();
                if (comp == null)
                    return false;

                SolidWorks.Interop.sldworks.MathUtility mathUtil =
                    (SolidWorks.Interop.sldworks.MathUtility)SolidWorksDocumentProvider._solidWorksApplication.GetMathUtility();

                // Transform the external point (assembly space) into the component's model space,
                // so it can be compared against the plane transforms which are in component model space.
                MathTransform compTransformInverse = (MathTransform)comp.Transform2.Inverse();
                MathPoint extPoint = (MathPoint)mathUtil.CreatePoint(externalPointCoords);
                extPoint = (MathPoint)extPoint.MultiplyTransform(compTransformInverse);
                double[] localCoords = (double[])extPoint.ArrayData;

                // Compute signed distances in component model space.
                // Opposite signs mean the external point is between the two planes.
                double distToMating = GetSignedDistanceToPlane(mathUtil, localCoords, matingPlane.Transform);
                double distToFree   = GetSignedDistanceToPlane(mathUtil, localCoords, freePlane.Transform);

                return distToMating * distToFree <= 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking shell intersection: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Returns the signed distance from a point to a plane.
        /// Positive = point is on the side the plane normal points toward.
        /// Negative = point is on the opposite side.
        /// </summary>
        private double GetSignedDistanceToPlane(
            SolidWorks.Interop.sldworks.MathUtility mathUtil,
            double[] point,
            MathTransform planeTransform)
        {
            // Transform (0,0,1) to get the plane normal in world space
            SolidWorks.Interop.sldworks.MathVector normal =
                (SolidWorks.Interop.sldworks.MathVector)mathUtil.CreateVector(new double[] { 0, 0, 1 });
            normal = (SolidWorks.Interop.sldworks.MathVector)normal.MultiplyTransform(planeTransform);
            double[] n = (double[])normal.ArrayData;

            // Transform (0,0,0) to get the plane origin in world space
            SolidWorks.Interop.sldworks.MathPoint origin =
                (SolidWorks.Interop.sldworks.MathPoint)mathUtil.CreatePoint(new double[] { 0, 0, 0 });
            origin = (SolidWorks.Interop.sldworks.MathPoint)origin.MultiplyTransform(planeTransform);
            double[] o = (double[])origin.ArrayData;

            // Signed distance = dot( (point - origin), normal )
            return (point[0] - o[0]) * n[0]
                 + (point[1] - o[1]) * n[1]
                 + (point[2] - o[2]) * n[2];
        }
    }
}

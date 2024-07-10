using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class InnerDishedEnd : DishedEnd
    {

        private const string INNER_DISHED_END_NAME = "Inner end";
        private const string POSITION_PLANE_NAME = "Position plane";
        private const string DISHED_END_ECOMPONENT_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\Inner dished end.SLDPRT";
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string CENTER_AXIS_NAME = "Center axis";

        /// <summary>
        /// Gets or sets the alignment of the dished end. If the alignment changes, a method to modify the existing alignment is triggered.
        /// </summary>
        public void SetAlignment(DishedEndAlignment dishedEndAlignment)
        {
            if (dishedEndAlignment != GetAlignment())
                ChangeAlignment();
        }

        /// <summary>
        /// DO NOT DELETE IT!!! This constructor is needed for Json deserializer.
        /// </summary>
        public InnerDishedEnd() { }

        /// <summary>
        /// This constructor is called when a new inner dished end is added.
        /// Creates and positions an inner dished end within a SolidWorks assembly.
        /// </summary>
        /// <param name="assemblyOfDishedEndsCenterAxis"></param>
        /// <param name="assemblyOfDishedEndsFrontPlane"></param>
        /// <param name="referenceDishedEnd"></param>
        /// <param name="dishedEndAlignment"></param>
        /// <param name="distance"></param>
        /// <param name="compartmentNumber"></param>
        public InnerDishedEnd(
            Feature assemblyOfDishedEndsCenterAxis,
            Feature assemblyOfDishedEndsFrontPlane,
            DishedEnd referenceDishedEnd,
            DishedEndAlignment dishedEndAlignment,
            double distance,
            int compartmentNumber)
            : base()
        {
            // 1. Input Validation
            if (assemblyOfDishedEndsCenterAxis == null)
                throw new ArgumentNullException(nameof(assemblyOfDishedEndsCenterAxis));

            if (assemblyOfDishedEndsFrontPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfDishedEndsFrontPlane));

            if (referenceDishedEnd == null)
                throw new ArgumentNullException(nameof(referenceDishedEnd));

            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.GetActiveDoc();
            if (assemblyOfDishedEndsDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            string positionPlaneName = $"{INNER_DISHED_END_NAME} {compartmentNumber} {POSITION_PLANE_NAME}";
            string componentName = $"{INNER_DISHED_END_NAME} {compartmentNumber}";

            // 2. Create Position Plane
            Feature positionPlane = FeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referenceDishedEnd.GetPositionPlane(),
                distance: distance,
                name: positionPlaneName);

            // 3. Add and Make Independent Dished End Component
            Component2 dishedEnd = ComponentManager.AddComponentPart(DISHED_END_ECOMPONENT_PATH);

            ComponentManager.MakeDishedEndIndependent(dishedEnd, DISHED_END_ECOMPONENT_PATH);

            // 4. Rename the Component
            FeatureManager.GetFeatureByName(SolidWorksDocumentProvider.GetActiveDoc(), dishedEnd.Name2).Name = componentName;

            // 5. Get Features for Mating
            Feature innerDishedEndCenterAxis = FeatureManager.GetFeatureByName(dishedEnd, "Center axis");

            // 6. Create Mates
            Feature rightPlaneMate = MateManager.CreateMate(
                componentFeature1: positionPlane,
                componentFeature2: FeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Right),
                alignmentType: MateAlignment.Aligned,
                name: $"{dishedEnd.Name2} - {RIGHT_PLANE_NAME}");

            Feature frontPlaneMate = MateManager.CreateMate(
                componentFeature1: assemblyOfDishedEndsFrontPlane,
                componentFeature2: FeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Front),
                alignmentType: MateAlignment.Aligned,
                name: $"{dishedEnd.Name2} - {FRONT_PLANE_NAME}");

            Feature centerAxisMate = MateManager.CreateMate(
                componentFeature1: assemblyOfDishedEndsCenterAxis,
                componentFeature2: innerDishedEndCenterAxis,
                alignmentType: MateAlignment.Anti_Aligned,
                name: $"{dishedEnd.Name2} - {CENTER_AXIS_NAME}");

            // 7. Get Inner Dished End Entities and Initialize Settings
            _dishedEndSettings = new DishedEndSettings();
            try
            {
                GetInnerDishedPIDs();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting inner dished end entities: {ex.Message}");
            }

            try
            {
                //Set the required position to the DishedEndPosition property
                SetAlignment(dishedEndAlignment);
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error changing inner dished end's alignment: {ex.Message}");
            }

            // 8. Helper Method to Get Inner Dished End Entities
            void GetInnerDishedPIDs()
            {
                try
                {
                    // Populate the _dishedEndSettings with the retrieved PIDs
                    _dishedEndSettings.PIDPositionPlane = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(positionPlane);
                    _dishedEndSettings.PIDComponent = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(dishedEnd);
                    _dishedEndSettings.PIDCenterAxis = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(innerDishedEndCenterAxis);
                    _dishedEndSettings.PIDCenterAxisMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(centerAxisMate);
                    _dishedEndSettings.PIDRightPlaneMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(rightPlaneMate);
                    _dishedEndSettings.PIDFrontPlaneMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(frontPlaneMate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "The attribute could not be created.");
                    return;
                }
            }
        }

        /// <summary>
        /// Gets whether dished end is alligned left or right.
        /// </summary>
        /// <returns></returns>
        private DishedEndAlignment GetAlignment()
        {
            //Get the transformation matrix from the dishedEnd object
            MathTransform transform = GetComponent().Transform2;

            //Transform the reference point (1, 0, 0) using the transformation matrix
            double[] TransformedVector = MathUtility.TransformVector(SolidWorksDocumentProvider._solidWorksApplication, transform, new double[3] { 1, 0, 0 });

            //Determine orientation based on the transformed point's X-coordinate
            return (TransformedVector[0] > 0 ? DishedEndAlignment.Left : DishedEndAlignment.Right);
        }

        /// <summary>
        /// Changes alignment of the dished end component
        /// </summary>
        public void ChangeAlignment()
        {
            try
            {
                FeatureManager.Suppress(GetCenterAxisMate());

                //Change alignment of the component
                //Warning message if ChangeAlignement() did not work
                if (!MateManager.ChangeAlignment(GetRightPlaneMate()))
                {
                    MessageBox.Show("Failed to change the right plane mate alignment.");
                    FeatureManager.Unsuppress(GetCenterAxisMate());
                    return;
                }

                //Change alignment of axis
                //Warning message if ChangeAlignement() did not work
                if (!MateManager.ChangeAlignment(GetCenterAxisMate()))
                {
                    MessageBox.Show("Failed to change the center axis mate alignment.");
                    MateManager.ChangeAlignment(GetRightPlaneMate());
                    FeatureManager.Unsuppress(GetCenterAxisMate());
                    return;
                }

                //Unsuppress axis mate
                FeatureManager.Unsuppress(GetCenterAxisMate());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Mate alignment could not be changed." + ex.Message);
                return;
            }
        }

        /// <summary>
        /// Deletes the dished end component
        /// </summary>
        public void Delete()
        {
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr selectionManager = (SelectionMgr)assemblyOfDishedEndsDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            GetComponent().Select4(false, selectData, false);
            GetPositionPlane().Select2(true, 1);

            //Get dished end document's path to delete the file
            ModelDoc2 componentDocument = GetComponent().GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected dished end
            ((AssemblyDoc)assemblyOfDishedEndsDoc).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            assemblyOfDishedEndsDoc.EditRebuild3();

            //Delete the file
            File.Delete(path);
        }
    }
}

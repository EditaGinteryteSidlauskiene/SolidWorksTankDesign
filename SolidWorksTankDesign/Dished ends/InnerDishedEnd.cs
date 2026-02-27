using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    public class InnerDishedEnd : DishedEnd
    {

        private const string INNER_DISHED_END_NAME = "Inner end";
        private const string POSITION_PLANE_NAME = "Position plane";
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string CENTER_AXIS_NAME = "Center axis";

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
            string projectFolder,
            string innerDishedEndDocPath,
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
            Feature positionPlane = SWFeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referenceDishedEnd.GetPositionPlane(),
                distance: distance,
                name: positionPlaneName,
                flip: false);

            // Create a path where the dished end doc will be saved
            string targetPath = Path.Combine(projectFolder,
                $"{INNER_DISHED_END_NAME} {SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds.Count+1}.SLDPRT");

            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;

            // Open dished end doc and save it to a new destination
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(innerDishedEndDocPath);
            ModelDoc2 innerDishedEndDoc = solidWorksApp.OpenDoc7(documentSpecification);

            innerDishedEndDoc.SaveAs3(targetPath, 0, 0);

            // Close dished end doc and newly saved docs
            solidWorksApp.CloseDoc(innerDishedEndDoc.GetTitle());
            solidWorksApp.CloseDoc(targetPath);

            // 3. Add and Make Independent Dished End Component
            Component2 dishedEnd = ComponentManager.AddComponentPart(targetPath);

            // 4. Rename the Component
            SWFeatureManager.GetFeatureByName(SolidWorksDocumentProvider.GetActiveDoc(), dishedEnd.Name2).Name = componentName;

            // 5. Get Features for Mating
            Feature innerDishedEndCenterAxis = SWFeatureManager.GetFeatureByName(dishedEnd, "Center axis");

            Feature rightPlaneMate = null;
            Feature frontPlaneMate = null;
            Feature centerAxisMate = null;

            try
            {
                // 6. Create Mates
                rightPlaneMate = MateManager.CreateMate(
                    componentFeature1: positionPlane,
                    componentFeature2: SWFeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{dishedEnd.Name2} - {RIGHT_PLANE_NAME}");

                frontPlaneMate = MateManager.CreateMate(
                    componentFeature1: assemblyOfDishedEndsFrontPlane,
                    componentFeature2: SWFeatureManager.GetMajorPlane(dishedEnd, MajorPlane.Front),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{dishedEnd.Name2} - {FRONT_PLANE_NAME}");

                centerAxisMate = MateManager.CreateMate(
                    componentFeature1: assemblyOfDishedEndsCenterAxis,
                    componentFeature2: innerDishedEndCenterAxis,
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{dishedEnd.Name2} - {CENTER_AXIS_NAME}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


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
                ComponentManager.SetAlignment(
                    dishedEnd,
                    dishedEndAlignment,
                    centerAxisMate,
                    rightPlaneMate,
                    _dishedEndSettings);
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

            assemblyOfDishedEndsDoc.EditRebuild3();
        }


        /// <summary>
        /// Deletes the dished end component
        /// </summary>
        public async void Delete()
        {
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr selectionManager = (SelectionMgr)assemblyOfDishedEndsDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            Feature rightPlaneMate = GetRightPlaneMate();
            Feature frontPlaneMate = GetFrontPlaneMate();
            Feature centerAxisMate = GetCenterAxisMate();
            Component2 dishedEndComponent = GetComponent();
            Feature dishedEndPostionPlane = GetPositionPlane();

            rightPlaneMate.Select2(false, 1);
            ((AssemblyDoc)assemblyOfDishedEndsDoc).DeleteSelections(1);

            frontPlaneMate.Select2(false, 1);
            ((AssemblyDoc)assemblyOfDishedEndsDoc).DeleteSelections(1);

            centerAxisMate.Select2(false, 1);
            ((AssemblyDoc)assemblyOfDishedEndsDoc).DeleteSelections(1);

            dishedEndComponent.Select4(false, selectData, false);
            dishedEndPostionPlane.Select2(true, 1);

            //Get dished end document's path to delete the file
            ModelDoc2 componentDocument = GetComponent().GetModelDoc2();
            string path = componentDocument.GetPathName();
            string docTitle = componentDocument.GetTitle();

            //Delete selected dished end
            ((AssemblyDoc)assemblyOfDishedEndsDoc).DeleteSelections(1);

            assemblyOfDishedEndsDoc.ClearSelection2(true);

            // 5) Close the document
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(docTitle);

            assemblyOfDishedEndsDoc.ForceRebuild3(true);

            SolidWorksDocumentProvider._filesToDelete.Add(path);
        }
    }
}

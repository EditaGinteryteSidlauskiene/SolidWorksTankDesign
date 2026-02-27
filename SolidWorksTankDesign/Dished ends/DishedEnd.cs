using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    public class DishedEnd
    {
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string CENTER_AXIS_NAME = "Center axis";

        public DishedEndSettings _dishedEndSettings;

        public DishedEnd()
        {
            _dishedEndSettings = new DishedEndSettings();
        }

        public void CompleteDishedEnd(
            string projectFolder,
            string dishedEndPath,
            Feature positionPlane,
            Feature assemblyOfDishedEndsCenterAxis,
            Feature assemblyOfDishedEndsFrontPlane,
            DishedEndAlignment dishedEndAlignment)
        {
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Create a path where the dished end doc will be saved
            string targetPath = Path.Combine(projectFolder,
                $"{dishedEndAlignment} dished end.SLDPRT");

            // Open dished end doc and save it to a new destination
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(dishedEndPath);
            ModelDoc2 emptyManholeDoc = solidWorksApp.OpenDoc7(documentSpecification);

            emptyManholeDoc.SaveAs3(targetPath, 0, 0);

            // Close dished end doc and newly saved docs
            solidWorksApp.CloseDoc(emptyManholeDoc.GetTitle());
            solidWorksApp.CloseDoc(targetPath);

            Component2 dishedEnd = ComponentManager.AddComponentPart(targetPath);

            // 5. Get Features for Mating
            Feature dishedEndCenterAxis = SWFeatureManager.GetFeatureByName(dishedEnd, "Center axis");

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
                    componentFeature2: dishedEndCenterAxis,
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{dishedEnd.Name2} - {CENTER_AXIS_NAME}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // 7. Get Main Dished Ends Entities and Initialize Settings
            if (dishedEndAlignment == DishedEndAlignment.Left)
                _dishedEndSettings = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings;
            else
                _dishedEndSettings = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings;

            try
            {
                GetDishedPIDs();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting inner dished end entities: {ex.Message}");
            }

            // 8. Helper Method to Get Inner Dished End Entities
            void GetDishedPIDs()
            {
                try
                {
                    // Populate the _dishedEndSettings with the retrieved PIDs
                    _dishedEndSettings.PIDPositionPlane = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(positionPlane);
                    _dishedEndSettings.PIDComponent = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(dishedEnd);
                    _dishedEndSettings.PIDCenterAxis = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(dishedEndCenterAxis);
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

            // Update documents and close assembly of dished ends
            DocumentManager.UpdateAndSaveDocuments();
        }

        public Feature GetCenterAxis() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxis,
                        out int error);

        public Feature GetPositionPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDPositionPlane,
                        out int error);

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDComponent,
                        out int error);

        public Feature GetCenterAxisMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxisMate,
                        out int error);

        public Feature GetRightPlaneMate() => (Feature) SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDRightPlaneMate,
                        out int error);

        public Feature GetFrontPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDFrontPlaneMate,
                        out int error);

        /// <summary>
        /// Repositions this dished end by updating its reference plane to match the 
        /// position plane of the specified new reference dished end.
        /// </summary>
        /// <param name="newReferenceDishedEnd"></param>
        public void RepositionByReference(DishedEnd newReferenceDishedEnd)
        {
            string name = newReferenceDishedEnd.GetComponent().Name2;
            // Validate the new reference dished end's position plane
            Feature newDishedEndPositionPlane = newReferenceDishedEnd.GetPositionPlane();
            if(newDishedEndPositionPlane == null)
            {
                MessageBox.Show($"Could not change reference plane, because could not find position plane of {newReferenceDishedEnd.GetComponent().Name2}.");
                return;
            }

            try
            {
                // Change the reference plane of this dished end
                bool success = SWFeatureManager.ChangeReferenceOfReferencePlane(
                    newDishedEndPositionPlane,  // New reference plane
                    GetPositionPlane());           // This dished end's current position plane

                //Warning message if ChangeReferenceOfReferencePlane() did not work
                if (!success)
                {
                    MessageBox.Show($"Could change reference plane of {GetComponent().Name2}.");
                    return;
                }

                ModelDoc2 activeDoc = SolidWorksDocumentProvider.GetActiveDoc();
                activeDoc.Save3(
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                            (int)swFileSaveError_e.swGenericSaveError,
                            (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        /// <summary>
        /// Changes dished end position plane's distance
        /// </summary>
        /// <param name="distance"></param>
        public void ChangeDistance(double distance)
        {
            SWFeatureManager.ChangeDistanceOfReferencePlane(GetPositionPlane(), distance);

            ComponentManager.RefreshDishedEnds();

            SolidWorksDocumentProvider.GetActiveDoc().Save3(
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                            (int)swFileSaveError_e.swGenericSaveError,
                            (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }
    }
}

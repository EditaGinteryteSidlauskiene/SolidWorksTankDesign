using SolidWorks.Interop.sldworks;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class DishedEnd
    {
        public DishedEndSettings _dishedEndSettings;

        public DishedEnd()
        {
            _dishedEndSettings = new DishedEndSettings();
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

        public Feature GetRightPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
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
                bool success = FeatureManager.ChangeReferenceOfReferencePlane(
                    SolidWorksDocumentProvider.GetActiveDoc(),
                    newDishedEndPositionPlane,  // New reference plane
                    GetPositionPlane());           // This dished end's current position plane

                //Warning message if ChangeReferenceOfReferencePlane() did not work
                if (!success)
                {
                    MessageBox.Show($"Could change reference plane of {GetComponent().Name2}.");
                    return;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

       
    }
}

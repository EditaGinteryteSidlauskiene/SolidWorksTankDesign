using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WarningAndErrorService;

namespace SolidWorksTankDesign
{
    internal class DishedEnd
    {
        public DishedEndSettings _dishedEndSettings;

        public DishedEnd()
        {
            _dishedEndSettings = new DishedEndSettings();
        }

        public Feature CenterAxis() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxis,
                        out int error);

        public Feature PositionPlane() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDPositionPlane,
                        out int error);

        public Component2 Component() => (Component2)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDComponent,
                        out int error);

        public Feature CenterAxisMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDCenterAxisMate,
                        out int error);

        public Feature RightPlaneMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDRightPlaneMate,
                        out int error);
        public Feature FrontPlaneMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _dishedEndSettings.PIDFrontPlaneMate,
                        out int error);

        /// <summary>
        /// Repositions this dished end by updating its reference plane to match the 
        /// position plane of the specified new reference dished end.
        /// </summary>
        /// <param name="newReferenceDishedEnd"></param>
        public void RepositionByReference(DishedEnd newReferenceDishedEnd)
        {
            string name = newReferenceDishedEnd.Component().Name2;
            // Validate the new reference dished end's position plane
            Feature newDishedEndPositionPlane = newReferenceDishedEnd.PositionPlane();
            if(newDishedEndPositionPlane == null)
            {
                MessageBox.Show($"Could not change reference plane, because could not find position plane of {newReferenceDishedEnd.Component().Name2}.");
                return;
            }

            try
            {
                // Change the reference plane of this dished end
                bool success = FeatureManager.ChangeReferenceOfReferencePlane(
                    SolidWorksDocumentProvider.ActiveDoc(),
                    newDishedEndPositionPlane,  // New reference plane
                    PositionPlane());           // This dished end's current position plane

                //Warning message if ChangeReferenceOfReferencePlane() did not work
                if (!success)
                {
                    MessageBox.Show($"Could change reference plane of {Component().Name2}.");
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

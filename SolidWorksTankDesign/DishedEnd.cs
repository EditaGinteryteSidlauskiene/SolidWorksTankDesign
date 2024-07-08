using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

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
    }
}

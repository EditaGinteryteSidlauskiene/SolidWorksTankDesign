using SolidWorks.Interop.sldworks;

namespace SolidWorksTankDesign
{
    internal class CylindricalShell
    {
        public CylindricalShellSettings _cylindricalShellSettings;

        public CylindricalShell() 
        {
            _cylindricalShellSettings = new CylindricalShellSettings();
        }

        public Feature GetLeftEndPlane() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDLeftEndPlane,
                        out int error);

        public Feature GetRightEndPlane() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDRightEndPlane,
                        out int error);

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDComponent,
                        out int error);

        public Feature GetCenterAxis() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDCenterAxis,
                        out int error);

        public Feature GetCenterAxisMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDCenterAxisMate,
                        out int error);

        public Feature GetLeftEndMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDLeftEndMate,
                        out int error);

        public Feature GetFrontPlaneMate() => (Feature)SolidWorksDocumentProvider.ActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDFrontPlaneMate,
                        out int error);
    }
}

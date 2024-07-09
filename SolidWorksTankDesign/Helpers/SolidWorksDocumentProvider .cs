using SolidWorks.Interop.sldworks;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// This class serves as a central point for managing and accessing both the active SolidWorks document (ModelDoc2) 
    /// and the SolidWorks application instance(SldWorks) within the project.
    /// </summary>
    internal static class SolidWorksDocumentProvider
    {
        public static SldWorks _solidWorksApplication;
        public static TankSiteAssembly _tankSiteAssembly;

        public static ModelDoc2 ActiveDoc()
        {
            return _solidWorksApplication.IActiveDoc2; 
        }
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.TankSiteConfigurations;
using System.Collections.Generic;

namespace SolidWorksTankDesign
{
    /// <summary>
    /// This class serves as a central point for managing and accessing both the active SolidWorks document (ModelDoc2) 
    /// and the SolidWorks application instance(SldWorks) within the project.
    /// </summary>
    public static class SolidWorksDocumentProvider
    {
        public static SldWorks _solidWorksApplication;
        public static TankSiteAssembly _tankSiteAssembly;
        public static TankProperties _tankProperties;
        public static List<string> _filesToDelete;
        // Central project folder path for the current project. Set once at startup/creation.
        public static string ProjectFolderPath { get; set; }

        public static ModelDoc2 GetActiveDoc()
        {
            return _solidWorksApplication.IActiveDoc2; 
        }
    }
}

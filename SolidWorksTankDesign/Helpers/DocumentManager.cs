using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal static  class DocumentManager
    {
        /// <summary>
        /// Updates tanksite assembly attribute, saves and closes subassembly, and saves tank site assembly doc.
        /// </summary>
        /// <param name="subassemblyModelDoc"></param>
        public static void UpdateAndSaveDocuments()
        {
            // Update SW attribute parameter
            TankSiteAssemblyDataManager.SerializeAndStoreTankSiteAssemblyData();

            //Save and close all subassemblies starting from the lowest in the hierarchy
            while (!ReferenceEquals(SolidWorksDocumentProvider._solidWorksApplication.ActiveDoc, SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc))
            {
                ModelDoc2 subassemblyDoc = SolidWorksDocumentProvider.GetActiveDoc();

                subassemblyDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(subassemblyDoc.GetTitle());
            }

            // Save the document of tank site assembly
            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }
    }
}

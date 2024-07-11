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
        public static void UpdateAndSaveDocuments(ModelDoc2 subassemblyModelDoc)
        {
            // Update SW attribute parameter
            TankSiteAssemblyDataManager.SerializeAndStoreTankSiteAssemblyData();

            // If the passed subassembly is currently active, save it and close
            if (SolidWorksDocumentProvider._solidWorksApplication.ActiveDoc() == subassemblyModelDoc)
            {
                // Save and the document of the passed subassembly
                subassemblyModelDoc.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    (int)swFileSaveError_e.swGenericSaveError,
                    (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(subassemblyModelDoc.GetTitle());
            }

            else
            {
                MessageBox.Show($"Cannot save {subassemblyModelDoc.GetTitle()} because it is not active");
            }
           
            // Save the document of tank site assembly
            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }
    }
}

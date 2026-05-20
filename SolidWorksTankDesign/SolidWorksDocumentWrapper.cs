using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;

namespace AddinWithTaskpane
{
    /// <summary>
    /// This class closes document automatically.
    /// </summary>
    internal class SolidWorksDocumentWrapper : IDisposable
    {
        private SldWorks solidWorksApplication;
        private ModelDoc2 modelDoc;

        public SolidWorksDocumentWrapper(SldWorks SolidWorksApplication, ModelDoc2 ModelDoc)
        {
            this.modelDoc = ModelDoc;
            this.solidWorksApplication = SolidWorksApplication;
            solidWorksApplication.ActivateDoc3(modelDoc.GetTitle(), true, 0, 0);
        }

        /// <summary>
        /// Closes document and releases the COM RCW to prevent DisconnectedContext errors.
        /// </summary>
        public void Dispose()
        {
            string title = modelDoc.GetTitle();
            solidWorksApplication.CloseDoc(title);
            Marshal.ReleaseComObject(modelDoc);
            modelDoc = null;
        }
    }
}

using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        /// Closes document
        /// </summary>
        public void Dispose()
        {
            solidWorksApplication.CloseDoc(modelDoc.GetTitle());
        }
    }
}

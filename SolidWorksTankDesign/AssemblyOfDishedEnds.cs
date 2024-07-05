using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WarningAndErrorService;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    //This is the highest class of dished ends. It contains all dished ends in the project
    internal class AssemblyOfDishedEnds
    {
        SldWorks _solidWorksApplication;
        private readonly ModelDoc2 _assemblyOfDishedEndsModelDoc;
        private readonly ModelDocExtension _assemblyOfDishedEndsModelDocExt;
        private readonly AssemblyDoc _assemblyOfDishedEndsAssemblyDoc;

        [JsonProperty("LeftDishedEnd")]
        public DishedEnd LeftDishedEnd { get; private set; }

        [JsonProperty("RightDishedEnd")]
        public DishedEnd RightDishedEnd { get; private set; }

        [JsonProperty("InnerDishedEnd")]
        public List<DishedEnd> InnerDishedEnds;

        public Feature CenterAxis
        {
            get
            {
                using (var document = new SolidWorksDocumentWrapper(_solidWorksApplication, _assemblyOfDishedEndsModelDoc))
                {
                    return FeatureManager.GetFeatureByName(_assemblyOfDishedEndsModelDoc, "Center axis");
                }
            }
        }

        public AssemblyOfDishedEnds() { }

        /// <summary>
        /// Initializes a TankSiteAssembly object, representing a SolidWorks tank site assembly model.It ensures 
        /// the provided SolidWorks application and model document are valid and performs initial setup operations. 
        /// </summary>
        /// <param name="warningService"></param>
        /// <param name="solidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public AssemblyOfDishedEnds(SldWorks solidWorksApplication, ModelDoc2 assemblyOfDisehdEndsModelDoc, Component2 leftDishedEndComponent, Component2 rightDishedEndComponent)
        {
            _solidWorksApplication = solidWorksApplication;

            if (assemblyOfDisehdEndsModelDoc == null)
            {
                throw new ArgumentNullException("Assembly of dished ends component is required.");
            }

            // Store references to the SolidWorks application and model objects
            _assemblyOfDishedEndsModelDoc = assemblyOfDisehdEndsModelDoc;

            // Get the model document extension for additional functionality
            _assemblyOfDishedEndsModelDocExt = _assemblyOfDishedEndsModelDoc.Extension;

            // Cast the model document to an AssemblyDoc
            _assemblyOfDishedEndsAssemblyDoc = (AssemblyDoc)_assemblyOfDishedEndsModelDoc;

            //---------------- Initialization -------------------------------------
            

            LeftDishedEnd = new DishedEnd(_solidWorksApplication, _assemblyOfDishedEndsModelDoc, leftDishedEndComponent);
            RightDishedEnd = new DishedEnd(_solidWorksApplication, _assemblyOfDishedEndsModelDoc, rightDishedEndComponent);
            
            try
            {
            }
            catch (Exception ex) { }
        }
    }
}

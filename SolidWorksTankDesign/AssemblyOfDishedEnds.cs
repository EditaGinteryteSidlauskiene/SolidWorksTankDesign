using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;

namespace SolidWorksTankDesign
{
    //This is the highest class of dished ends. It contains all dished ends in the project
    internal class AssemblyOfDishedEnds
    {
        [JsonProperty("LeftDishedEnd")]
        public DishedEnd LeftDishedEnd { get; private set; }

        [JsonProperty("RightDishedEnd")]
        public DishedEnd RightDishedEnd { get; private set; }

        [JsonProperty("InnerDishedEnd")]
        public List<InnerDishedEnd> InnerDishedEnds = new List<InnerDishedEnd>();

        public Feature CenterAxis() => FeatureManager.GetFeatureByName(SolidWorksDocumentProvider.ActiveDoc(), "Center axis");

        public AssemblyOfDishedEnds()
        {
        }

        /// <summary>
        /// Initializes a TankSiteAssembly object, representing a SolidWorks tank site assembly model.It ensures 
        /// the provided SolidWorks application and model document are valid and performs initial setup operations. 
        /// </summary>
        /// <param name="warningService"></param>
        /// <param name="solidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public AssemblyOfDishedEnds(ModelDoc2 assemblyOfDisehdEndsModelDoc)
        {
            if (assemblyOfDisehdEndsModelDoc == null)
            {
                throw new ArgumentNullException("Assembly of dished ends component is required.");
            }

            //---------------- Initialization -------------------------------------

            LeftDishedEnd = new DishedEnd();
            RightDishedEnd = new DishedEnd();
        }

        /// <summary>
        /// Creates a new InnerDishedEnd object, and adds it to the list
        /// </summary>
        /// <param name="referenceDishedEnd"></param>
        /// <param name="dishedEndAlignment"></param>
        /// <param name="distance"></param>
        /// <param name="compartmentNumber"></param>
        public void AddInnerDishedEnd(DishedEnd referenceDishedEnd, DishedEndAlignment dishedEndAlignment, double distance, int compartmentNumber)
        {
            InnerDishedEnds.Add(
                new InnerDishedEnd(
                CenterAxis(),
                FeatureManager.GetMajorPlane(SolidWorksDocumentProvider.ActiveDoc(), MajorPlane.Front),
                referenceDishedEnd,
                dishedEndAlignment,
                distance,
                compartmentNumber)
                );
        }
    }
}

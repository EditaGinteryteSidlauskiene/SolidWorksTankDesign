using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    //This is the highest class of dished ends. It contains all dished ends in the project
    internal class AssemblyOfDishedEnds
    {
        private ModelDoc2 currentlyActiveDishedEndsDoc;

        [JsonProperty("LeftDishedEnd")]
        public DishedEnd LeftDishedEnd { get; private set; }

        [JsonProperty("RightDishedEnd")]
        public DishedEnd RightDishedEnd { get; private set; }

        [JsonProperty("InnerDishedEnd")]
        public List<InnerDishedEnd> InnerDishedEnds = new List<InnerDishedEnd>();

        public AssemblyOfDishedEnds() { }

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
        /// Removes the last dished end component from the assembly.
        /// </summary>
        private bool RemoveInnerEnd()
        {
            int innerEndsCount = InnerDishedEnds.Count;
            // Check if there are any dished ends to remove
            if (innerEndsCount == 0)
            {
                MessageBox.Show("There are no inner dished ends in the list.");
                return false;
            }

            // Delete the last dished end (assuming it controls the SolidWorks object)
            InnerDishedEnds[innerEndsCount - 1].Delete();

            // Remove the last dished end reference from the tracking list
            InnerDishedEnds.RemoveAt(innerEndsCount - 1);

            return true;
        }

        public Feature GetCenterAxis() => FeatureManager.GetFeatureByName(currentlyActiveDishedEndsDoc, "Center axis");
    
        /// <summary>
        /// Activates document of assembly of dished ends
        /// </summary>
        public void ActivateDocument()
        {
            ModelDoc2 AssemblyOfDishedEndsModelDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetDishedEndsAssemblyComponent().GetModelDoc2();
            currentlyActiveDishedEndsDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(AssemblyOfDishedEndsModelDoc.GetTitle()+ ".sldasm", true, 0, 0);
        }

        /// <summary>
        /// Closes active document of assembly of dished ends
        /// </summary>
        public void CloseDocument()
        {
            if (currentlyActiveDishedEndsDoc == null) return;

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(currentlyActiveDishedEndsDoc.GetTitle());
            currentlyActiveDishedEndsDoc = null;
        }

        /// <summary>
        /// Creates a new InnerDishedEnd object, and adds it to the list
        /// After this method, call TankSiteAssembly.SerializeAndStoreTankAssemblyData()!!!
        /// </summary>
        /// <param name="referenceDishedEnd"></param>
        /// <param name="dishedEndAlignment"></param>
        /// <param name="distance"></param>
        /// <param name="compartmentNumber"></param>
        public void AddInnerDishedEnd(DishedEnd referenceDishedEnd, DishedEndAlignment dishedEndAlignment, double distance, int compartmentNumber)
        {
            try
            {
                InnerDishedEnds.Add(
                new InnerDishedEnd(
                GetCenterAxis(),
                FeatureManager.GetMajorPlane(SolidWorksDocumentProvider.GetActiveDoc(), MajorPlane.Front),
                referenceDishedEnd,
                dishedEndAlignment,
                distance,
                compartmentNumber)
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Sets the number of inner dished ends in the assembly, adding or removing them 
        /// as needed to match the required count. Adjusts the position of the rightmost 
        /// dished end based on the final configuration.
        /// </summary>
        /// <param name="requiredNumberOfDishedEnds">The desired number of inner dished ends.</param>
        /// <param name="defaultDishedEndAlignment">The default alignment for newly added dished ends.</param>
        /// <param name="defaultDistance">The default distance between dished ends.</param>
        public void SetNumberOfInnerDishedEnds(int requiredNumberOfDishedEnds, DishedEndAlignment defaultDishedEndAlignment, double defaultDistance)
        {
            ActivateDocument();

            // --- 1. Handle Cases Where Fewer Dished Ends Are Needed ---

            if (requiredNumberOfDishedEnds < InnerDishedEnds.Count)
            {
                // Reposition the rightmost dished end to maintain continuity.
                // If there are no inner dished ends left, use the left dished end as the reference.
                RightDishedEnd.RepositionByReference(requiredNumberOfDishedEnds == 0 ? LeftDishedEnd : InnerDishedEnds[requiredNumberOfDishedEnds - 1]);

                // Remove excess dished ends until the count matches the required number.
                while (requiredNumberOfDishedEnds != InnerDishedEnds.Count)
                {
                    if (!RemoveInnerEnd()) return;
                }
            }

            // --- 2. Handle Cases Where More Dished Ends Are Needed ---

            else if (requiredNumberOfDishedEnds > InnerDishedEnds.Count)
            {
                // Add new dished ends until the count matches the required number.
                while (requiredNumberOfDishedEnds != InnerDishedEnds.Count)
                {
                    try
                    {
                        // Add a new inner dished end, using the previous one as a reference.
                        // If this is the first inner dished end, use the left dished end as the reference.
                        AddInnerDishedEnd(
                            (InnerDishedEnds.Count == 0 ? LeftDishedEnd : InnerDishedEnds[InnerDishedEnds.Count - 1]),
                            defaultDishedEndAlignment,
                            defaultDistance,
                            InnerDishedEnds.Count + 1);  // Use the next compartment number
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }

                // Reposition the rightmost dished end based on the final state of the inner dished ends.
                RightDishedEnd.RepositionByReference(InnerDishedEnds.Count == 0 ? LeftDishedEnd : InnerDishedEnds[InnerDishedEnds.Count - 1]);
            }

            else
            {
                CloseDocument();
                return;
            }

            // Update documents and close assembly of dished ends
            DocumentManager.UpdateAndSaveDocuments(currentlyActiveDishedEndsDoc);
            currentlyActiveDishedEndsDoc = null;
        }
    }
}

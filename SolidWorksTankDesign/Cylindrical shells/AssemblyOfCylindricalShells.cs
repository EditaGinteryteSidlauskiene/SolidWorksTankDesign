using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class AssemblyOfCylindricalShells
    {
        private ModelDoc2 currentlyActiveCylindricalShellsDoc;

        [JsonProperty("CylindricalShells")]
        public List<CylindricalShell> CylindricalShells = new List<CylindricalShell>();

        public Feature GetCenterAxis() => FeatureManager.GetFeatureByName(currentlyActiveCylindricalShellsDoc, "Center axis");

        public AssemblyOfCylindricalShells() { }

        /// <summary>
        /// Removes the last cyldrical shell component from the assembly.
        /// </summary>
        private bool RemoveCylindricalShell()
        {
            int cylindricalShellsCount = CylindricalShells.Count;
            // Check if there are any dished ends to remove

            if (cylindricalShellsCount == 0)
            {
                MessageBox.Show("There are no cylindrical shells in the assembly.");
                return false;
            }
            if (cylindricalShellsCount == 1)
            {
                MessageBox.Show("There is only one cylindrical shell in the assembly, and it cannot be removed.");
                return false;
            }

            // Delete the last dished end (assuming it controls the SolidWorks object)
            CylindricalShells[cylindricalShellsCount - 1].Delete();

            // Remove the last dished end reference from the tracking list
            CylindricalShells.RemoveAt(cylindricalShellsCount - 1);

            return true;
        }

        private void AddCylindricalShell(double length, double diameter)
        {
            ActivateDocument();
            
            try
            {
                CylindricalShells.Add(
                    new CylindricalShell(
                        CylindricalShells.Last(),
                        GetCenterAxis(),
                        FeatureManager.GetMajorPlane(SolidWorksDocumentProvider.GetActiveDoc(), MajorPlane.Front),
                        length,
                        diameter,
                        CylindricalShells.Count + 1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        public void SetNumberOfCylindricalShells(int requiredNumberOfCylindricalShells, double defaultLength, double diameter)
        {
            ActivateDocument();

            // --- 1. Handle Cases Where Fewer Cylindrical Shells Are Needed ---

            if (requiredNumberOfCylindricalShells < CylindricalShells.Count)
            {
                // Remove excess cylindrical shells until the count matches the required number.
                while (requiredNumberOfCylindricalShells != CylindricalShells.Count)
                {
                    if (!RemoveCylindricalShell()) return;
                }
            }

            else
            {
                // Add new cylindrical shells until the count matches the required number.
                while (requiredNumberOfCylindricalShells != CylindricalShells.Count)
                {
                    try
                    {
                        // Add a new cylindrical shell, using the previous one as a reference.
                        AddCylindricalShell(defaultLength, diameter);  
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }
            }

            // Update SW attribute parameter
            TankSiteAssemblyDataManager.SerializeAndStoreTankSiteAssemblyData();

            // Save and the document of assembly of dished ends
            currentlyActiveCylindricalShellsDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            CloseDocument();

            // Save the document of tank site assembly
            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        /// <summary>
        /// Activates document of assembly of cylindrical shells
        /// </summary>
        public void ActivateDocument()
        {
            ModelDoc2 AssemblyOfDCylindricalShellsModelDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetCylindricalShellsAssemblyComponent().GetModelDoc2();
            currentlyActiveCylindricalShellsDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(AssemblyOfDCylindricalShellsModelDoc.GetTitle() + ".sldasm", true, 0, 0);
        }

        /// <summary>
        /// Closes active document of assembly ofcylindrical shells
        /// </summary>
        public void CloseDocument()
        {
            if (currentlyActiveCylindricalShellsDoc == null) return;

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(currentlyActiveCylindricalShellsDoc.GetTitle());
            currentlyActiveCylindricalShellsDoc = null;
        }
    }
}

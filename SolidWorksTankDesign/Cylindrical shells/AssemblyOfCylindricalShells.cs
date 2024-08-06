using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class AssemblyOfCylindricalShells
    {
        private ModelDoc2 currentlyActiveCylindricalShellsDoc;

        [JsonProperty("CylindricalShells")]
        public List<CylindricalShell> CylindricalShells = new List<CylindricalShell>();

        public Feature GetCenterAxis() => SWFeatureManager.GetFeatureByName(currentlyActiveCylindricalShellsDoc, "Center axis");

        public AssemblyOfCylindricalShells() { }

        /// <summary>
        /// Removes the last cyldrical shell component from the assembly.
        /// </summary>
        private bool RemoveCylindricalShell()
        {
            // Get the count of cylindrical shells once for efficiency
            int cylindricalShellsCount = CylindricalShells.Count;

            // Check if there are enough cylindrical shells to remove
            if (cylindricalShellsCount == 0)
            {
                MessageBox.Show("There are no cylindrical shells in the assembly.");
                CloseDocument();
                return false;
            }
            if (cylindricalShellsCount == 1)
            {
                MessageBox.Show("There is only one cylindrical shell in the assembly, and it cannot be removed.");
                CloseDocument();
                return false;
            }

            try
            {
                // Attempt to delete the SolidWorks object associated with the shell
                CylindricalShells[cylindricalShellsCount - 1].Delete();
            }
            catch (Exception ex)
            {
                // Handle potential exceptions during deletion
                MessageBox.Show($"Error deleting cylindrical shell: {ex.Message}");
                return false;
            }

            // Remove the shell from the internal tracking list
            CylindricalShells.RemoveAt(cylindricalShellsCount - 1);

            return true;
        }

        /// <summary>
        /// Adds a new cylindrical shell to a SolidWorks assembly.
        /// </summary>
        /// <param name="length"></param>
        /// <param name="diameter"></param>
        private void AddCylindricalShell(double length, double diameter)
        {
            try
            {
                // Create and add the new cylindrical shell
                CylindricalShells.Add(
                    new CylindricalShell(
                        CylindricalShells.Last(),
                        GetCenterAxis(),
                        SWFeatureManager.GetMajorPlane(SolidWorksDocumentProvider.GetActiveDoc(), MajorPlane.Front),
                        length,
                        diameter,
                        CylindricalShells.Count + 1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Sets the number of cylindrical shells in the assembly, adding or removing them 
        /// as needed to match the required count.
        /// </summary>
        /// <param name="requiredNumberOfCylindricalShells"></param>
        /// <param name="defaultLength"></param>
        /// <param name="diameter"></param>
        public void SetNumberOfCylindricalShells(int requiredNumberOfCylindricalShells, double defaultLength, double diameter)
        {
            // Ensure the correct SolidWorks document is active for modification
            ActivateDocument();

            // --- 1. Handle Cases Where Fewer Cylindrical Shells Are Needed ---

            if(requiredNumberOfCylindricalShells == 0)
            {
                MessageBox.Show("At least 1 cylindrical shell must be left.");
                CloseDocument();
            }

            else if (requiredNumberOfCylindricalShells < CylindricalShells.Count)
            {
                // Remove excess cylindrical shells until the count matches the required number.
                while (requiredNumberOfCylindricalShells != CylindricalShells.Count)
                {
                    try
                    {
                        if (!RemoveCylindricalShell()) return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }
            }

            // --- 2. Handle cases where more cylindrical shells are needed ---

            else if(requiredNumberOfCylindricalShells > CylindricalShells.Count)
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
                        CloseDocument();
                        return;
                    }
                }
            }

            else
            {
                CloseDocument();
                return;
            }

            // Update documents and close assembly of cylindrical shell
            DocumentManager.UpdateAndSaveDocuments();
            currentlyActiveCylindricalShellsDoc = null;
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

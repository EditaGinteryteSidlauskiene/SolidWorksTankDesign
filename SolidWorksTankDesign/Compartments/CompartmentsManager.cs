using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class CompartmentsManager
    {
        private ModelDoc2 _currentlyActiveShellDoc;

        [JsonProperty("Compartments")]
        public List<Compartment> Compartments = new List<Compartment>();

        public Feature GetCenterAxis() => FeatureManager.GetFeatureByName(_currentlyActiveShellDoc, "Center axis");

        public CompartmentsManager() { }

        /// <summary>
        /// CURRENTLY ACTIVE SHELL DOC MUST BE ACTIVATED!!!
        /// </summary>
        /// <param name="dishedEnd"></param>
        /// <returns></returns>
        private Feature GetDishedEndPositionPlane(DishedEnd dishedEnd)
        {
            _currentlyActiveShellDoc = SolidWorksDocumentProvider.GetActiveDoc();
            //Left end plane@Assembly of Dished ends-1@Shell
            // Get position plane's name from assembly of dished ends perspective
            string positionPlaneName = GetPositionPlaneName(dishedEnd);

            SelectionMgr selectionMgr = _currentlyActiveShellDoc.SelectionManager;

            // Select position plane using this name
            _currentlyActiveShellDoc.Extension.SelectByID2(positionPlaneName, "PLANE", 0, 0, 0, false, 0, null, 0);

            // Get selected object
            return selectionMgr.GetSelectedObject6(1, -1);

            string GetPositionPlaneName(DishedEnd dishedEndObject)
            {
                // 1. Get assembly of dished ends document
                Component2 assemblyOfDishedEndsComp = SolidWorksDocumentProvider._tankSiteAssembly.GetDishedEndsAssemblyComponent();
                ModelDoc2 assemblyOfDishedEndsDoc = assemblyOfDishedEndsComp.GetModelDoc2();

                // 2. Activate assembly of dished ends doc to be able to get position plane
                SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfDishedEndsDoc.GetTitle(), true, 0, 0);

                // Get shell component name
                string shellCompFullName = SolidWorksDocumentProvider._tankSiteAssembly.GetShellAssemblyComponent().Name2;
                string[] shellNameParts = shellCompFullName.Split('/');
                string shellName = shellNameParts[shellNameParts.Length-1].Split('-')[0];

                // Get dished end's component name
                string dishedEndsCompFullName = assemblyOfDishedEndsComp.Name2;
                string[] dishedEndNameParts = dishedEndsCompFullName.Split('/');
                string dishedEndsCompName = dishedEndNameParts[dishedEndNameParts.Length-1];

                // 3. Get position plane
                Feature positionPlaneOfDishedEnd = dishedEndObject.GetPositionPlane();

                // Close assembly of dished ends doc
                SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfDishedEndsDoc.GetTitle());

                // 4. Get position plane's name from assembly of dished ends doc perspective
                return $"{positionPlaneOfDishedEnd.Name}@{dishedEndsCompName}@{shellName}";
            }
        }

        /// <summary>
        /// Adds a new compartment to a SolidWorks shell assembly.
        /// </summary>
        /// <param name="dishedEnd"></param>
        private void AddCompartment(DishedEnd dishedEnd, double distanceBetweenNozzleAndRefPlane)
        {
            try
            {
                // Create and add the new compartment
                Compartments.Add(
                    new Compartment(
                        FeatureManager.GetMajorPlane(_currentlyActiveShellDoc, MajorPlane.Front),
                        GetCenterAxis(),
                        GetDishedEndPositionPlane(dishedEnd),
                        Compartments.Count,
                        distanceBetweenNozzleAndRefPlane));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Removes the last compartment component from the shell assembly.
        /// </summary>
        private bool RemoveCompartment()
        {
            // Get the count of compartments once for efficiency
            int compartmentsCount = Compartments.Count;

            // Check if there are enough compartments to remove
            if (compartmentsCount == 0)
            {
                MessageBox.Show("There are no compartments in the shell assembly.");
                CloseDocument();
                return false;
            }
            if (compartmentsCount == 1)
            {
                MessageBox.Show("There is only one compartment in the shell assembly, and it cannot be removed.");
                CloseDocument();
                return false;
            }

            try
            {
                // Attempt to delete the SolidWorks object associated with the shell
                Compartments[compartmentsCount - 1].Delete();
            }
            catch (Exception ex)
            {
                // Handle potential exceptions during deletion
                MessageBox.Show($"Error deleting compartment: {ex.Message}");
                return false;
            }

            // Remove the compartment from the internal tracking list
            Compartments.RemoveAt(compartmentsCount - 1);

            return true;
        }

        /// <summary>
        /// Sets the number of compartments in the assembly, adding or removing them 
        /// as needed to match the required count.
        /// </summary>
        /// <param name="requiredNumberOfCompartments"></param>
        /// <param name="dishedEnd"></param>
        public void SetNumberOfCompartments(int requiredNumberOfCompartments, DishedEnd dishedEnd, double distanceBetweenNozzleAndRefPlane)
        {
            // Ensure the correct SolidWorks document is active for modification
            ActivateDocument();

            // --- 1. Handle Cases Where Fewer Compartments Are Needed ---

            if (requiredNumberOfCompartments == 0)
            {
                MessageBox.Show("At least 1 compartment must be left.");
                CloseDocument();
            }

            else if (requiredNumberOfCompartments < Compartments.Count)
            {
                // Remove excess compartments until the count matches the required number.
                while (requiredNumberOfCompartments != Compartments.Count)
                {
                    try
                    {
                        if (!RemoveCompartment()) return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }
            }

            // --- 2. Handle cases where more compartments are needed ---

            else if (requiredNumberOfCompartments > Compartments.Count)
            {
                // Add new compartments until the count matches the required number.
                while (requiredNumberOfCompartments != Compartments.Count)
                {
                    try
                    {
                        // Add a new compartment
                        AddCompartment(dishedEnd, distanceBetweenNozzleAndRefPlane);
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

            // Update documents and close assembly of shell
            DocumentManager.UpdateAndSaveDocuments(_currentlyActiveShellDoc);
            _currentlyActiveShellDoc = null;
        }

        /// <summary>
        /// Activates document of assembly of cylindrical shells
        /// </summary>
        public void ActivateDocument()
        {
            ModelDoc2 shellModelDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetShellAssemblyComponent().GetModelDoc2();
            _currentlyActiveShellDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(shellModelDoc.GetTitle() + ".sldasm", true, 0, 0);
        }

        /// <summary>
        /// Closes active document of assembly ofcylindrical shells
        /// </summary>
        public void CloseDocument()
        {
            if (_currentlyActiveShellDoc == null) return;

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(_currentlyActiveShellDoc.GetTitle());
            _currentlyActiveShellDoc = null;
        }
    }
}

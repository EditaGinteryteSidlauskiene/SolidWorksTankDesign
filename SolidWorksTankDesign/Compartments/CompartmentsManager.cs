using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class CompartmentsManager
    {
        private ModelDoc2 _currentlyActiveShellDoc;

        [JsonProperty("Compartments")]
        public List<Compartment> Compartments = new List<Compartment>();

        public Feature GetCenterAxis() => SWFeatureManager.GetFeatureByName(_currentlyActiveShellDoc, "Center axis");

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
                        SWFeatureManager.GetMajorPlane(_currentlyActiveShellDoc, MajorPlane.Front),
                        GetCenterAxis(),
                        GetDishedEndPositionPlane(dishedEnd),
                        Compartments.Count));

                Compartments.Last().ActivateDocument();
                Compartments.Last().AddNozzle(
                    Compartments.Count(),
                    Compartments.Last().GetLeftEndPlane(),
                    distanceBetweenNozzleAndRefPlane);
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

        public void SelectAllCylindricalShells()
        {
            _currentlyActiveShellDoc.ClearSelection2(true);

            int cylindricalShellsCount = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells.Count;

            for (int i = 0; i < cylindricalShellsCount; i++)
            {
                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();
                ModelDoc2 cylindricalShellsDoc = SolidWorksDocumentProvider.GetActiveDoc();

                Component2 cylindricalShell = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[i].GetComponent();

                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CloseDocument();

                cylindricalShell.Select2(true, 1);
            }
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
            DocumentManager.UpdateAndSaveDocuments();
            _currentlyActiveShellDoc = null;
        }

        /// <summary>
        /// Swaps the positions of two compartments within a SolidWorks assembly, 
        /// adjusting their associated mates and names to maintain model consistency.
        /// </summary>
        /// <param name="compartment1Number"></param>
        /// <param name="compartment2Number"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void SwapCompartments(int compartment1Number, int compartment2Number)
        {
            // Validate compartment numbers
            if (compartment1Number < 1 || compartment1Number > Compartments.Count ||
                compartment2Number < 1 || compartment2Number > Compartments.Count)
            {
                throw new ArgumentOutOfRangeException("Invalid compartment numbers.");
            }

            // Activate shell doc
            ActivateDocument();

            // Get compartments (adjusting for zero-based indexing)
            Compartment compartment1 = Compartments[compartment1Number - 1];
            Compartment compartment2 = Compartments[compartment2Number - 1];

            // Get the specific mate features and associated names for compartment 1:
            // - Left-end mate (used for the primary connection between compartments)
            // - Front plane mate (used for renaming)
            // - Center axis mate (used for renaming)
            Feature compartment1Mate = compartment1.GetLeftEndMate();
            Feature comp1FrontPlaneMate = compartment1.GetFrontPlaneMate();
            Feature comp1CenterAxisMate = compartment1.GetCenterAxisMate();
            string mate1Name = compartment1Mate.Name;
            string frontPlaneMateName1 = comp1FrontPlaneMate.Name;
            string centerAxisMateName1 = comp1CenterAxisMate.Name;

            // Store the index of the compartment in the collection (needed for later swapping).
            int index1 = compartment1Number - 1;

            // Get the specific mate features and associated names for compartment 2
            // (mirroring the process done for compartment 1)
            Feature compartment2Mate = compartment2.GetLeftEndMate();
            Feature comp2FrontPlaneMate = compartment2.GetFrontPlaneMate();
            Feature comp2CenterAxisMate = compartment2.GetCenterAxisMate();
            string mate2Name = compartment2Mate.Name;
            string frontPlaneMateName2 = comp2FrontPlaneMate.Name;
            string centerAxisMateName2 = comp2CenterAxisMate.Name;
            int index2 = compartment2Number - 1;

            try
            {
                // Edit the coincident mates that connect the compartments.
                // The EditCoincidentMate method is responsible for modifying the existing mates to establish the new relationships between compartments after they are swapped.
                // Each compartment's left-end mate is adjusted to reference the dished end position plane of the other compartment and the right plane of its own compartment.
                MateManager.EditCoincidentMate(
                    compartment1.GetLeftEndMate(),
                    compartment2.GetDishedEndPositionPlane(),
                    SWFeatureManager.GetMajorPlane((Component2)compartment1.GetComponent(), MajorPlane.Right));
                MateManager.EditCoincidentMate(
                    compartment2.GetLeftEndMate(),
                    compartment1.GetDishedEndPositionPlane(),
                    SWFeatureManager.GetMajorPlane((Component2)compartment2.GetComponent(), MajorPlane.Right));
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error trying to edit the mates.", ex.Message);
            }

            // These variables will store the modified mate features after the swapping process, which will be used later for renaming.
            Feature editedMate1 = null;
            Feature editedMate2 = null;

            try
            {
                // Switch the positions of the two compartments within the 'Compartments' collection
                // and update their associated properties, such as the persistent IDs (PIDs) of the dished end position planes.
                // This ensures that the data referencing the compartments remains correct after the swap.
                SwapObjectsAndProperties();

                // Update the persistent IDs (PIDs) of the mates that were edited in the previous steps.
                // These PIDs are used to uniquely identify the features within the SolidWorks document.
                UpdateMatesPIDs();

                // After the mates have been edited and the compartments swapped, their names and the names of their associated mates are updated.
                // This renaming step helps maintain consistency in the model's structure after the swap operation.
                RenameMatesAndComponents();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error trying to swap compartments", ex.Message);
            }
            

            ///Swaps the positions of two compartments within a collection and ensures that their associated references to dished end position 
            ///planes (identified by PIDs) remain correctly linked.
            void SwapObjectsAndProperties()
            {
                // 1. Store Dished End Position Plane References:
                // Retrieve the Persistent IDs (PIDs) of the dished end position planes for both compartments.
                // These PIDs are unique identifiers for the planes within the SolidWorks document and will be used to reassign them after the compartment positions are swapped.
                byte[] comp1DishedEndPositionPlanePID = compartment1._compartmentSettings.PIDDishedEndPositionPlane;
                byte[] comp2DishedEndPositionPlanePID = compartment2._compartmentSettings.PIDDishedEndPositionPlane;

                // 2. Swap Compartment Positions:
                // Remove both compartment objects from the 'Compartments' collection at their original indices.
                // This prepares for inserting them back into the collection in the reversed order.
                Compartments.RemoveAt(index2);  // Remove compartment 2 first (to avoid index issues)
                Compartments.RemoveAt(index1);

                // 3. Re-insert Compartments in Swapped Order:
                // Insert the compartments back into the collection at the swapped indices. 
                // Now, compartment2 will be at index1 and compartment1 will be at index2.
                Compartments.Insert(index1, compartment2);
                Compartments.Insert(index2, compartment1);

                // 4. Re-assign Dished End Position Plane PIDs:
                // Assign the stored PIDs of the dished end position planes back to the correct compartments after the swap.
                // This ensures that the plane references are correctly associated with the compartments in their new positions.
                Compartments[index1]._compartmentSettings.PIDDishedEndPositionPlane = comp1DishedEndPositionPlanePID;
                Compartments[index2]._compartmentSettings.PIDDishedEndPositionPlane = comp2DishedEndPositionPlanePID;
            }

            /// Updates the persistent IDs (PIDs) of two edited mates within a SolidWorks document. 
            /// PIDs are unique identifiers for features, and updating them ensures that the compartment objects 
            /// correctly reference the modified mates after they have been swapped in the assembly.
            void UpdateMatesPIDs()
            {
                // 1. Get Selection Manager:
                // Obtain the SelectionMgr object from the currently active SolidWorks document. 
                // This object is used to manage and interact with the selection of entities within the document.
                SelectionMgr selectionMgr = (SelectionMgr)_currentlyActiveShellDoc.SelectionManager;

                // 2. Get the First Edited Mate:
                // Select the first edited mate (identified by its name stored in 'mate1Name') using the document's SelectByID2 method.
                // This method selects an entity within the document based on its name and type ("MATE" in this case).
                _currentlyActiveShellDoc.Extension.SelectByID2(
                               mate1Name,
                               "MATE",
                               0, 0, 0,
                               false,
                               0, null, 0);
                // Retrieve the selected mate object from the selection manager. 
                // The GetSelectedObject6 method is used to get a specific object from the current selection set.
                editedMate1 = selectionMgr.GetSelectedObject6(1, -1);

                // Update the Persistent ID (PID) of the left-end mate in the compartment's settings to reference the newly edited mate.
                // PIDs are unique identifiers for features within the document, ensuring accurate references even after modifications.
                Compartments[index2]._compartmentSettings.PIDLeftEndMate = _currentlyActiveShellDoc.Extension.GetPersistReference3(editedMate1);

                // 3. Get the Second Edited Mate:
                // Repeat the process to select and retrieve the second edited mate (identified by 'mate2Name').
                _currentlyActiveShellDoc.Extension.SelectByID2(
                               mate2Name,
                               "MATE",
                               0, 0, 0,
                               false,
                               0, null, 0);
                editedMate2 = selectionMgr.GetSelectedObject6(1, -1);

                // Update the PID of the left-end mate for the other compartment.
                Compartments[index1]._compartmentSettings.PIDLeftEndMate = _currentlyActiveShellDoc.Extension.GetPersistReference3(editedMate2);
            }

            /// Ensures correct naming after swapping two compartments in a SolidWorks assembly. 
            /// It does this by temporarily renaming components and mates to avoid conflicts, 
            /// then swapping the names to match the new positions of the compartments in the assembly.
            void RenameMatesAndComponents()
            {
                // 1. Temporarily Rename Components:
                // Assign temporary names to the components involved in the swap.
                // This is done to avoid potential naming conflicts when swapping the actual names later.
                Compartments[index1].GetComponent().Name2 = "TemporaryCompartment1Name";
                Compartments[index2].GetComponent().Name2 = "TemporaryCompartment2Name";

                // 2. Swap Component Names:
                // Assign the original names of the components back to them, but in swapped order.
                // Since the components are now at different indices in the 'Compartments' collection, their original names will be correctly associated with the swapped positions.
                Compartments[index1].GetComponent().Name2 = compartment1.GetComponent().Name2;
                Compartments[index2].GetComponent().Name2 = compartment2.GetComponent().Name2;

                // 3. Temporarily Rename Mates:
                // Assign temporary names to the mates associated with the compartments.
                // This is similar to the temporary renaming of components and serves the same purpose of avoiding conflicts during the final name swap.
                editedMate1.Name = "Temporary dished end postion plane mate name1";
                comp1FrontPlaneMate.Name = "Temporary front plane mate name 1";
                comp1CenterAxisMate.Name = "Temporary center axis mate name 1";

                editedMate2.Name = "Temporary name2";
                comp2FrontPlaneMate.Name = "Temporary front plane mate name 2";
                comp2CenterAxisMate.Name = "Temporary center axis mate name 2";

                // 4. Swap Mate Names:
                // Assign the original names of the mates back to them, but in swapped order.
                // This step mirrors the swapping of component names, ensuring that the mate names align with the new positions of the compartments.
                editedMate1.Name = compartment2Mate.Name;       // Assign mate2's original name to mate1
                comp1FrontPlaneMate.Name = frontPlaneMateName2;
                comp1CenterAxisMate.Name = centerAxisMateName2;

                editedMate2.Name = compartment1Mate.Name;       // Assign mate1's original name to mate2
                comp2FrontPlaneMate.Name = frontPlaneMateName1;
                comp2CenterAxisMate.Name = centerAxisMateName1;
            }

            // Update and save the SolidWorks documents to reflect the changes made in the model
            DocumentManager.UpdateAndSaveDocuments();
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

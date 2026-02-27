using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Path = System.IO.Path;

namespace SolidWorksTankDesign.MVP.Models
{
    public class CompartmentConfigurationModel : ICompartmentConfigurationModel
    {
        private const string INNER_DISHED_END_DOC_NAME = "Inner dished end.SLDPRT";
        private const string DISHED_END_MAIN_DOC_NAME = "Main dished end.SLDPRT";
        private const string CYLINDRICAL_SHELL_DOC_NAME = "Cylindrical shell.SLDPRT";
        private const string DISHED_END_VOLUME_INCLUDE_PROP_NAME = "volume include";
        private const string DISHED_END_VOLUME_EXCLUDE_PROP_NAME = "volume exclude";
        private const string CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME = "volume per meter";

        public string _mainFolderPath;
        public string _compartmentDocPath;
        public string _projectFolder;


        private double _mainDishedEndIncludeVolume;
        private double _innerDishedEndExcludeVolume;
        private double _innerDishedEndIncludeVolume;
        private double _cylindricalShellVolumePerMeter;

        private bool _hasChanges = false;

        public ObservableCollection<CompartmentConfiguration> CompartmentConfigurations { get; } = new ObservableCollection<CompartmentConfiguration>();

        public ObservableCollection<Treatment> InternalTreatments { get; } = new ObservableCollection<Treatment>();

        public CompartmentConfigurationModel(string projectFolder, string mainFolderPath, string compartmentDocPath)
        {
            _mainFolderPath = mainFolderPath;
            _compartmentDocPath = compartmentDocPath;
            _projectFolder = projectFolder;

            // Get custom properties values
            _mainDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(DISHED_END_MAIN_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _innerDishedEndExcludeVolume = -SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_EXCLUDE_PROP_NAME);
            _innerDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _cylindricalShellVolumePerMeter = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(CYLINDRICAL_SHELL_DOC_NAME), CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME);

            GetCompartmentConfigurations();
            GetInternalTreatments();

            CompartmentsMappingHelper.SaveOriginalConfigurations();
        }

        /// <summary>
        /// Gets the path for a document.
        /// Gets the folder name which depends on instalation type, diameter, and class.
        /// Gets a folder's, which is in the main folder, where empty documents are stored, path.
        /// If the folder exists, adds document's name to folder's path and returns full document path
        /// </summary>
        /// <param name="documentName"></param>
        /// <returns></returns>
        private string GetDocumentPath(string documentName)
        {
            TankProperties tankProperties = SolidWorksDocumentProvider._tankProperties;
            string instalation;

            if (tankProperties.ConstructionStandard == ConstructionStandard.EN_12285_1)
                instalation = "U";

            else
                instalation = "A";

            // Get name folder where empty documents are stored
            string folderName = Path.Combine(
                $"{instalation}T",
                $"{instalation}T {tankProperties.NominalDiameter} Class {tankProperties.Class}");

            // Get the path of that folder. It is stored in the main folder
            string folderPath = Path.Combine(_mainFolderPath, folderName);

            // If the folder exists
            if (Directory.Exists(folderPath))
            {
                // Return the full document path
                return Path.Combine(folderPath, documentName);
            }

            // Else, message that the folder was not found
            else
            {
                MessageBox.Show($"Folder {folderName} was not found in {_mainFolderPath} path.");
                return string.Empty;
            }
        }

        /// <summary>
        /// Adjusts the length of the cylindrical shell in the assembly based on the total length
        /// of all compartmentConfig configurations. Handles document activation and updates the model.
        /// </summary>
        private void ChangeCylindricalShellLength()
        {
            // Retrieve the assembly of cylindrical shells.
            AssemblyOfCylindricalShells assemblyOfCylindricalShells = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells;
            if(assemblyOfCylindricalShells == null)
            {
                MessageBox.Show("Could not change length of cylindrical shell.");
                return;
            }

            // Calculate the total length of all compartments.
            List<CompartmentConfiguration> compartmentsConfigurations = SolidWorksDocumentProvider._tankProperties?.CompartmentsConfigurations;
            if (compartmentsConfigurations == null || !compartmentsConfigurations.Any())
            {
                MessageBox.Show("Could not change length of cylindrical shell. Compartment Configurations are null or empty.");
                return;
            }

            double totalCylindricalShellLength = 0;
            foreach (CompartmentConfiguration compartmentConfiguration in compartmentsConfigurations)
            {
                totalCylindricalShellLength += ConvertMillimetersToMeters(compartmentConfiguration.Length); // Convert mm to meters.
            }

            try
            {
                // Activate the document to make changes.
                assemblyOfCylindricalShells.ActivateDocument();

                // Change the length of the first cylindrical shell.
                if (assemblyOfCylindricalShells.CylindricalShells.Any())
                {
                    assemblyOfCylindricalShells.CylindricalShells[0].ChangeLength(totalCylindricalShellLength);
                }
                else
                {
                    MessageBox.Show("Could not change length of cylindrical shell. No cylindrical shells were found.");
                }

                // Update and save the changes.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not change length of cylindrical shell: {ex.Message}");
                return;
            }
        }

        /// <summary>
        /// Converts length from millimeters to meters.
        /// </summary>
        /// <param name="lengthInMillimeters">The length in millimeters.</param>
        /// <returns>The length in meters.</returns>
        private double ConvertMillimetersToMeters(double lengthInMillimeters)
        {
            return lengthInMillimeters / 1000;
        }

        /// <summary>
        /// Updates the lengths of specified compartments based on their configurations.
        /// Activates the document for each compartmentConfig, modifies its length, and saves changes.
        /// </summary>
        /// <param name="compartmentsToModifyLengths">
        /// A list of tuples containing compartments and their corresponding configurations.
        /// </param>
        private void ChangeCompartmentsLengths(List<(Compartment compartment, CompartmentConfiguration compartmentConfiguration)> compartmentsToModifyLengths)
        {
            if (compartmentsToModifyLengths == null || !compartmentsToModifyLengths.Any())
            {
                MessageBox.Show("Could not change length of compartments. The list of compartments to modify is null or empty.");
                return;
            }

            foreach ((Compartment compartment, CompartmentConfiguration compartmentConfiguration) in compartmentsToModifyLengths)
            {
                if (compartment == null || compartmentConfiguration == null)
                {
                    MessageBox.Show($"Could not change length of one of the compartments.");
                    continue;
                }

                try
                {
                    // Activate the document for the compartmentConfig.
                    compartment.ActivateDocument();

                    // Adjust the compartmentConfig length (convert mm to meters).
                    double newLength = ConvertMillimetersToMeters(compartmentConfiguration.Length);
                    compartment.ChangeLength(newLength);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not change length of the compartmentConfig {compartmentConfiguration.Name}: {ex.Message}");
                }
            }

            try
            {
                // Perform a single save operation after all modifications.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving documents: {ex.Message}");
            }
        }

        /// <summary>
        /// Finds the appropriate dished end for a given compartmentConfig.
        /// </summary>
        /// <param name="compartment">The current compartmentConfig.</param>
        /// <param name="compartments">The list of all compartments.</param>
        /// <param name="assemblyOfDishedEnds">The assembly containing dished ends.</param>
        /// <returns>The corresponding dished end, or null if not found.</returns>
        private DishedEnd GetDishedEndForCompartment(Compartment compartment)
        {
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;
            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds;

            int index = FindCompartmentIndex(compartment);

            if(index == -1) return null;

            // If the compartmentConfig is the last one, return right dished end, else inner dished end
            return index != compartments.Count - 1
                ? assemblyOfDishedEnds.InnerDishedEnds[index]
                : assemblyOfDishedEnds.RightDishedEnd;
        }

        /// <summary>
        /// Updates the distances for the dished ends based on the specified compartmentConfig configurations.
        /// </summary>
        /// <param name="compartmentsToModifyLengths">
        /// A list of tuples containing compartments and their corresponding configurations.
        /// </param>
        private void ChangeDishedEndDistance(List<(Compartment compartment, CompartmentConfiguration compartmentConfiguration)> compartmentsToModifyLengths)
        {
            if (compartmentsToModifyLengths == null || !compartmentsToModifyLengths.Any())
            {
                MessageBox.Show($"Could not change distances of dished ends. The list of compartments to modify is null or empty.");
                return;
            }

            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly?._assemblyOfDishedEnds;
            if (assemblyOfDishedEnds == null)
            {
                MessageBox.Show($"Could not change distances of dished ends. Assembly of dished ends is null.");
                return;
            }

            // Activate the assembly of dished ends document.
            assemblyOfDishedEnds.ActivateDocument();

            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly?._compartmentsManager?.Compartments;
            if (compartments == null)
            {
                MessageBox.Show($"Could not change distances of dished ends. Compartments list is null.");
                return;
            }

            foreach ((Compartment compartment, CompartmentConfiguration compartmentConfiguration) in compartmentsToModifyLengths)
            {
                if (compartment == null || compartmentConfiguration == null)
                {
                    MessageBox.Show($"Could not change distance of the dished end next to {compartmentConfiguration.Name}. " +
                        $"The compartmentConfig or its configuration is null.");
                    continue;
                }

                DishedEnd dishedEnd = null;
                try
                {
                    // Find the corresponding dished end.
                    dishedEnd = GetDishedEndForCompartment(compartment);
                    if (dishedEnd != null)
                    {
                        // Change the distance (convert mm to meters).
                        double newDistance = ConvertMillimetersToMeters(compartmentConfiguration.Length);
                        dishedEnd.ChangeDistance(newDistance);
                    }
                    else
                    {
                        MessageBox.Show($"Could not change distance of the dished end {dishedEnd.GetComponent().Name2}. The dished end was not found.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error changing distance of the dished end {dishedEnd.GetComponent().Name2}: {ex.Message}");
                }
            }

            try
            {
                // Ensure all modifications are saved.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving documents: {ex.Message}");
            }
        }

        /// <summary>
        /// Finds the index of a given compartmentConfig in the list of compartments.
        /// </summary>
        /// <param name="compartment">The compartmentConfig to find.</param>
        /// <param name="compartments">The list of compartments.</param>
        /// <returns>The index of the compartmentConfig, or -1 if not found.</returns>
        private int FindCompartmentIndex(Compartment compartment)
        {
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;

            for (int i = 0; i < compartments.Count; i++)
            {
                if (compartments[i]._compartmentSettings.ID == compartment._compartmentSettings.ID)
                {
                    return i;
                }
            }
            return -1; // Not found
        }

        /// <summary>
        /// Determines if the alignment of the left dished end has changed for the specified compartmentConfig
        /// compared to its configuration.
        /// </summary>
        /// <param name="compartment">The compartmentConfig to check.</param>
        /// <param name="compartmentConfiguration">The configuration containing the expected alignment.</param>
        /// <returns>True if the alignment has changed, otherwise false.</returns>
        private bool LeftDishedEndAlignmentChanged(Compartment compartment, CompartmentConfiguration compartmentConfiguration)
        {
            if (compartment == null || compartmentConfiguration == null)
            {
                return false;
            }

            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly?._compartmentsManager?.Compartments;
            List<InnerDishedEnd> innerDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly?._assemblyOfDishedEnds?.InnerDishedEnds;

            if (compartments == null || innerDishedEnds == null)
            {
                MessageBox.Show($"Cannot perform alignment check. Lists of Compartments or InnerDishedEnds are null.");
                return false;
            }

            // Find the index of the specified compartmentConfig.
            int index = FindCompartmentIndex(compartment);
            if (index < 0 || index > innerDishedEnds.Count)
            {
                MessageBox.Show($"Compartment {compartmentConfiguration.Name} is not valid or out of range.");
                return false;
            }

            DishedEnd dishedEnd = null;
            try
            {
                // Get the dished end corresponding to the compartmentConfig.
                dishedEnd = index > 0 ? innerDishedEnds[index - 1] : SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd;

                DishedEndAlignment currentAlignment = MateManager.GetAlignment(dishedEnd);

                // Compare the current alignment with the expected alignment.
                if (compartmentConfiguration.LeftDishedEndAlignment != currentAlignment)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error checking alignment for dished end for compartmentConfig {compartmentConfiguration.Name}: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Retrieves the inner dished end for the specified compartmentConfig.
        /// </summary>
        /// <param name="compartment">The compartmentConfig to find the dished end for.</param>
        /// <param name="compartments">The list of all compartments.</param>
        /// <param name="assemblyOfDishedEnds">The assembly containing dished ends.</param>
        /// <returns>The corresponding inner dished end, or null if not found.</returns>
        private InnerDishedEnd GetInnerDishedEndForCompartment(Compartment compartment)
        {
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager?.Compartments;
            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds;

            for (int i = 0; i < compartments.Count; i++)
            {
                if (compartments[i]._compartmentSettings.ID == compartment._compartmentSettings.ID)
                {
                    return i > 0 ? assemblyOfDishedEnds.InnerDishedEnds[i - 1] : null;
                }
            }
            return null; // Not found
        }

        /// <summary>
        /// Changes the alignments of left dished ends for the specified compartments based on their configurations.
        /// </summary>
        /// <param name="compartmentsToModifyLeftDishedAlignments">
        /// A list of tuples containing compartments and their configurations.
        /// </param>
        private void ChangeLeftDishedEndsAlignments(List<(Compartment compartment, CompartmentConfiguration compartmentConfiguration)> compartmentsToModifyLeftDishedAlignments)
        {
            if (compartmentsToModifyLeftDishedAlignments == null || !compartmentsToModifyLeftDishedAlignments.Any())
            {
                MessageBox.Show($"Cannot change alignments of dished ends. The list of compartments to modify is null or empty.");
                return;
            }

            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly?._assemblyOfDishedEnds;
            if (assemblyOfDishedEnds == null)
            {
                MessageBox.Show($"Cannot change alignments of dished ends. Assembly of dished ends is null.");
                return;
            }

            ModelDoc2 dishedEndsDoc = assemblyOfDishedEnds.ActivateDocument();
            if (dishedEndsDoc == null)
            {
                MessageBox.Show($"Cannot change alignments of dished ends. Assembly of dished ends is null. Failed to activate the assembly of dished ends document.");
                return;
            }

            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly?._compartmentsManager?.Compartments;
            if (compartments == null || !compartments.Any())
            {
                MessageBox.Show($"Cannot change alignments of dished ends. List of Compartments is null or empty.");
                return;
            }

            foreach (var compartmentPair in compartmentsToModifyLeftDishedAlignments)
            {
                if (compartmentPair.compartment == null || compartmentPair.compartmentConfiguration == null)
                {
                    MessageBox.Show($"One of inner disheds alignment could not be changed. A compartmentConfig or its configuration is null.");
                    continue;
                }

                DishedEnd innerDishedEnd = null;
                try
                {
                    // Get the corresponding dished end.
                    innerDishedEnd = GetInnerDishedEndForCompartment(compartmentPair.compartment);
                    if (innerDishedEnd == null)
                    {
                        MessageBox.Show($"Cannot change alignment of dished end for compartmentConfig {compartmentPair.compartmentConfiguration.Name}." +
                            $"Inner dished end not found.");
                        continue;
                    }

                    // Change the alignment of the dished end.
                    ComponentManager.ChangeAlignment(
                        innerDishedEnd.GetCenterAxisMate(),
                        innerDishedEnd.GetRightPlaneMate(),
                        innerDishedEnd._dishedEndSettings);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error changing alignment for compartmentConfig {innerDishedEnd.GetComponent().Name2}: {ex.Message}");
                }
            }

            try
            {
                // Save changes after all modifications are complete.
                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving documents: {ex.Message}");
            }
        }

        /// <summary>
        /// The method generates a unique name for a new compartmentConfig by determining the next 
        /// alphabetical letter based on existing compartmentConfig names in the CompartmentConfigurations list.
        /// </summary>
        /// <returns></returns>
        private string GetCompartmentName()
        {
            char latestLetter = 'A';

            foreach (CompartmentConfiguration config in CompartmentConfigurations)
            {
                string[] nameParts = config.Name.Split(' ');

                // Ensure the name follows the expected format
                if (nameParts.Length > 1 && !string.IsNullOrWhiteSpace(nameParts[1]))
                {
                    char letter = nameParts[1][0];
                    if (char.IsLetter(letter) && letter > latestLetter)
                    {
                        latestLetter = letter;
                    }
                }
            }

            // Increment the letter for the new compartmentConfig name
            return $"Compartment {(char)(latestLetter + 1)}";
        }

        /// <summary>
        /// Finds the index of a compartment by its ID
        /// </summary>
        /// <param name="compartments"></param>
        /// <param name="id"></param>
        /// <returns></returns>
        private int FindCompartmentIndexById(List<Compartment> compartments, Guid id)
        {
            return compartments.FindIndex(comp => comp._compartmentSettings.ID == id);
        }

        /// <summary>
        /// Adjusts the tank assembly's dished ends after a compartment is deleted. 
        /// It ensures that: 
        /// - Remaining dished ends are repositioned to maintain structural alignment.
        /// - Inner dished ends corresponding to the deleted compartment are removed. 
        /// - Configuration properties are updated for the affected compartments.
        /// </summary>
        /// <param name="deletedCompartmentIndex"></param>
        private void UpdateDishedEndAfterCompartmentDeletion(int deletedCompartmentIndex)
        {
            // Retrieve the current list of compartments and the assembly of dished ends.
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;
            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds;
            assemblyOfDishedEnds.ActivateDocument();

            // Initialize variables for compartment mapping and dished end adjustments.
            (Compartment compartment, CompartmentConfiguration config)? mapping = null;

            // Retrieve the configuration mapping of the previous compartment, if applicable.
            if (deletedCompartmentIndex > 0)
                mapping = CompartmentsMappingHelper.GetMappingById(
                SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[deletedCompartmentIndex - 1]._compartmentSettings.ID);

            // Variables for the dished end to reposition, its reference, and the inner dished end to delete.
            DishedEnd dishedEndToReposition;
            DishedEnd referenceDishedEnd;
            InnerDishedEnd innerDishedEndToDelete;
            int innerDishedEndToDeleteIndex;

            // Case 1: Deleting the first compartment
            if (deletedCompartmentIndex == 0)
            {
                // Determine the dished end to reposition and reference for alignment.
                dishedEndToReposition = compartments.Count == 1 ? assemblyOfDishedEnds.RightDishedEnd : assemblyOfDishedEnds.InnerDishedEnds[1];
                referenceDishedEnd = assemblyOfDishedEnds.LeftDishedEnd;
                innerDishedEndToDeleteIndex = 0;
                innerDishedEndToDelete = assemblyOfDishedEnds.InnerDishedEnds[0];

                // Update the configuration for the first compartment after deletion.
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[0].LeftDishedEndAlignment = DishedEndAlignment.Left;
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[0].LeftEndConnection = LeftEndConnection.None;
            }

            // Case 2: Deleting a middle compartment
            else if (deletedCompartmentIndex != compartments.Count)
            {
                // Determine the reference dished end for alignment.
                int referenceDishedEndIndex = deletedCompartmentIndex - 2;
                dishedEndToReposition = assemblyOfDishedEnds.InnerDishedEnds[deletedCompartmentIndex];
                referenceDishedEnd = referenceDishedEndIndex >= 0 ?
                    assemblyOfDishedEnds.InnerDishedEnds[referenceDishedEndIndex] : assemblyOfDishedEnds.LeftDishedEnd;
                innerDishedEndToDeleteIndex = deletedCompartmentIndex - 1;
                innerDishedEndToDelete = assemblyOfDishedEnds.InnerDishedEnds[deletedCompartmentIndex - 1];

                // Adjust the distance of the repositioned dished end based on the configuration.
                assemblyOfDishedEnds.InnerDishedEnds[deletedCompartmentIndex].ChangeDistance(mapping.Value.config.Length / 1000);
            }

            // Case 3: Deleting the last compartment when only one remains
            else if (deletedCompartmentIndex == compartments.Count && compartments.Count == 1)
            {
                dishedEndToReposition = assemblyOfDishedEnds.RightDishedEnd;
                referenceDishedEnd = assemblyOfDishedEnds.LeftDishedEnd;
                innerDishedEndToDeleteIndex = deletedCompartmentIndex - 1;
                innerDishedEndToDelete = assemblyOfDishedEnds.InnerDishedEnds[innerDishedEndToDeleteIndex];

                // Adjust the distance of the right dished end based on the configuration.
                assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(
                   mapping.Value.config.Length / 1000);
            }

            // Case 4: Deleting the last compartment with more than one compartment present
            else
            {
                dishedEndToReposition = assemblyOfDishedEnds.RightDishedEnd;
                referenceDishedEnd = assemblyOfDishedEnds.InnerDishedEnds[deletedCompartmentIndex - 2];
                innerDishedEndToDeleteIndex = deletedCompartmentIndex - 1;
                innerDishedEndToDelete = assemblyOfDishedEnds.InnerDishedEnds[innerDishedEndToDeleteIndex];

                // Adjust the distance of the right dished end based on the configuration.
                assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(
                    mapping.Value.config.Length / 1000);
            }

            // Reposition the target dished end based on its reference dished end.
            dishedEndToReposition.RepositionByReference(referenceDishedEnd);

            // Delete the inner dished end corresponding to the deleted compartment.
            innerDishedEndToDelete.Delete();

            // Remove the deleted inner dished end from the list.
            assemblyOfDishedEnds.InnerDishedEnds.RemoveAt(innerDishedEndToDeleteIndex);
        }

        /// <summary>
        /// Deletes a compartment and updates dependent features
        /// </summary>
        /// <param name="compartments"></param>
        /// <param name="index"></param>
        private void DeleteCompartment(List<Compartment> compartments, int index)
        {
            // Case: Deleting the first compartment
            if (index == 0)
            {
                // Save the document before making changes
                DocumentManager.UpdateAndSaveDocuments();

                // Activate the document containing the compartment assembly
                ModelDoc2 shellDoc = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();

                // Retrieve the component and mate information for the next compartment
                Component2 nextCompartmentComp = compartments[1].GetComponent();
                Feature mateToEdit = compartments[1].GetLeftEndMate();

                // Get the new position plane for the left dished end and the right plane of the next compartment
                Feature newDishedEndPositionPlane = SWFeatureManager.GetDishedEndPositionPlane(SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd);
                Feature compartmentRightPlane = SWFeatureManager.GetMajorPlane(nextCompartmentComp, MajorPlane.Right);

                // Edit the mate to align the next compartment with the left dished end
                MateManager.EditCoincidentMate(mateToEdit, newDishedEndPositionPlane, compartmentRightPlane);

                // Update the mate's PID to ensure it is properly tracked in the document
                SWFeatureManager.UpdateMatePID(shellDoc, SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[1], $"{mateToEdit.Name}");

                // Save the document after making changes
                DocumentManager.UpdateAndSaveDocuments();
            }

            // Delete the compartment at the specified index
            compartments[index].Delete();

            // Remove the compartment from the list to maintain consistency
            compartments.RemoveAt(index);

            // Update dished ends and mates after deleting the compartment
            UpdateDishedEndAfterCompartmentDeletion(index);

            // Save the document to persist all changes
            DocumentManager.UpdateAndSaveDocuments();
        }


        /// <summary>
        /// Handles the deletion of compartments based on their IDs
        /// </summary>
        /// <param name="compartments"></param>
        /// <param name="deletedIds"></param>
        private void HandleDeletedCompartments(List<Compartment> compartments, List<Guid> deletedIds)
        {
            foreach (Guid deletedId in deletedIds)
            {
                int index = FindCompartmentIndexById(compartments, deletedId);
                if (index >= 0)
                {
                    DeleteCompartment(compartments, index);
                }
            }
        }

        /// <summary>
        /// The ModifyTank method manages modifications to the tank compartments in a SolidWorks assembly.
        /// It handles deleted, added, and switched compartment configurations while ensuring that all dependent features
        /// (e.g., mates, dished ends, and positions) are updated. The method also saves updated configurations
        /// and rebuilds the SolidWorks document.
        /// </summary>
        public void ModifyTank()
        {
            // Apply changes in compartment configurations (changes of length and left dished end alignment)
            // before handling major operations (deleting, adding and swapping)
            ApplyCompartmentsChanges();

            // Retrieve relevant data
            List<Compartment> compartments = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments;
            List<Guid> deletedCompartmentsIds = CompartmentsMappingHelper.GetDeletedConfigurations();
            List<(CompartmentConfiguration config, int index)> addedCompartmentConfigs = CompartmentsMappingHelper.GetAddedConfigurations();
            List<(CompartmentConfiguration config, int index)> switchedConfigs = CompartmentsMappingHelper.GetSwitchedConfigurations(CompartmentConfigurations);

            // Handle deleted compartments
            if (deletedCompartmentsIds.Count > 0)
            {
                _hasChanges = true;
                HandleDeletedCompartments(compartments, deletedCompartmentsIds);
            }

            
            if (addedCompartmentConfigs.Count > 0)
            {
                _hasChanges = true;
                foreach ((CompartmentConfiguration config, int index) addedCompartmentConfig in addedCompartmentConfigs)
                {
                    List<CompartmentConfiguration> compartmentConfigurations = SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations;
                    // 1. Add inner dished end
                    AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds;
                    ModelDoc2 assemblyOfDishedEndsDoc = assemblyOfDishedEnds.ActivateDocument();

                    // Get reference dished end
                    DishedEnd referenceDishedEnd = null;
                    double length = 0;
                    DishedEndAlignment alignment = DishedEndAlignment.Left;
                    int innerDishedEndPosition = 0;

                    // if new compartmentConfig config is the last one of two, get the LeftDishedEnd
                    if (addedCompartmentConfig.index == SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Count - 1 &&
                        SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Count == 2)
                    {
                        referenceDishedEnd = assemblyOfDishedEnds.LeftDishedEnd;
                        length = compartmentConfigurations[addedCompartmentConfig.index - 1].Length / 1000;
                        alignment = compartmentConfigurations[addedCompartmentConfig.index].LeftDishedEndAlignment;
                        innerDishedEndPosition = addedCompartmentConfig.index - 1;
                    }

                    // If the new compartmentConfig config is the last one of more than 2
                    else if (addedCompartmentConfig.index == SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Count - 1)
                    {
                        referenceDishedEnd = assemblyOfDishedEnds.InnerDishedEnds.Last();
                        length = compartmentConfigurations[addedCompartmentConfig.index - 1].Length / 1000;
                        alignment = compartmentConfigurations[addedCompartmentConfig.index].LeftDishedEndAlignment;
                        innerDishedEndPosition = addedCompartmentConfig.index - 1;
                    }

                    // If the new compartmentConfig config is the first one
                    else if (addedCompartmentConfig.index == 0)
                    {
                        referenceDishedEnd = assemblyOfDishedEnds.LeftDishedEnd;
                        length = addedCompartmentConfig.config.Length / 1000;
                        alignment = compartmentConfigurations[1].LeftDishedEndAlignment;
                        innerDishedEndPosition = 0;
                    }

                    // If the new compartmentConfig confing is the second one
                    else if (addedCompartmentConfig.index == 1)
                    {
                        referenceDishedEnd = assemblyOfDishedEnds.LeftDishedEnd;
                        length = compartmentConfigurations[0].Length / 1000;
                        alignment = compartmentConfigurations[1].LeftDishedEndAlignment;
                        innerDishedEndPosition = 0;
                    }

                    // Else if the new compartmentConfig config is anywhere else (neither first, second, nor last)
                    else
                    {
                        referenceDishedEnd = assemblyOfDishedEnds.InnerDishedEnds[addedCompartmentConfig.index - 2];
                        length = compartmentConfigurations[addedCompartmentConfig.index - 1].Length / 1000;
                        alignment = compartmentConfigurations[addedCompartmentConfig.index].LeftDishedEndAlignment;
                        innerDishedEndPosition = addedCompartmentConfig.index - 1;
                    }

                    // Add inner dished end. It will be positioned correctly in SolidWorks, but it will be the last one in
                    // SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds list
                    assemblyOfDishedEnds.AddInnerDishedEnd(
                        _projectFolder,
                        GetDocumentPath(INNER_DISHED_END_DOC_NAME),
                        referenceDishedEnd,
                        alignment,
                        length,
                        addedCompartmentConfig.index);

                    // Reposition inner dished end in SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds list
                    InnerDishedEnd addedInnerDishedEnd = assemblyOfDishedEnds.InnerDishedEnds.Last();
                    assemblyOfDishedEnds.InnerDishedEnds.RemoveAt(assemblyOfDishedEnds.InnerDishedEnds.Count - 1);
                    assemblyOfDishedEnds.InnerDishedEnds.Insert(innerDishedEndPosition, addedInnerDishedEnd);

                    DishedEnd compartmentDishedEndReference = addedInnerDishedEnd;
                    CompartmentsManager compartmentsManager = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager;

                    // 2. Move next inner dished end or right dished end
                    //If the added inner dished end is the last one, move right dished end and change distance
                    if (addedCompartmentConfig.index == assemblyOfDishedEnds.InnerDishedEnds.Count)
                    {
                        assemblyOfDishedEnds.RightDishedEnd.RepositionByReference(
                            addedInnerDishedEnd);

                        assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(addedCompartmentConfig.config.Length / 1000);

                        SolidWorksDocumentProvider.GetActiveDoc().EditRebuild3();
                    }

                    // Else, move next inner dished end and change distance if newly added compartmentConfig confing is NOT the first one
                    else if (addedCompartmentConfig.index != 0)
                    {
                        assemblyOfDishedEnds.InnerDishedEnds[addedCompartmentConfig.index].RepositionByReference(
                            addedInnerDishedEnd);

                        assemblyOfDishedEnds.InnerDishedEnds[addedCompartmentConfig.index].ChangeDistance(addedCompartmentConfig.config.Length / 1000);

                        SolidWorksDocumentProvider.GetActiveDoc().EditRebuild3();
                    }

                    // If the added compartmentConfig confing is the first one, move the second inner dished end
                    else
                    {
                        if (compartmentsManager.Compartments.Count == 1) assemblyOfDishedEnds.RightDishedEnd.RepositionByReference(
                           addedInnerDishedEnd);

                        else assemblyOfDishedEnds.InnerDishedEnds[1].RepositionByReference(
                           addedInnerDishedEnd);

                        compartmentDishedEndReference = assemblyOfDishedEnds.LeftDishedEnd;
                        Feature innerDishedEndPositionPlane = SWFeatureManager.GetDishedEndPositionPlane(addedInnerDishedEnd);

                        ModelDoc2 shellDoc = compartmentsManager.ActivateDocument();

                        ModelDoc2 compartmentModelDoc = compartmentsManager.Compartments[0].GetComponent().GetModelDoc2();
                        Feature compartmentRightMajorPlane = null;

                        compartmentRightMajorPlane = SWFeatureManager.GetMajorPlane(compartmentsManager.Compartments[0].GetComponent(), MajorPlane.Right);

                        Feature compartmentLeftEndMate = compartmentsManager.Compartments[0].GetLeftEndMate();
                        // Move the compartmentConfig next to the one to be added (currently, the first one)
                        MateManager.EditCoincidentMate(
                            compartmentLeftEndMate,
                            innerDishedEndPositionPlane,
                            compartmentRightMajorPlane);

                        Feature updatedMate = MateManager.GetMateByName(shellDoc, $"{compartmentLeftEndMate.Name}");
                        compartmentsManager.Compartments[0]._compartmentSettings.PIDLeftEndMate = shellDoc.Extension.GetPersistReference3(updatedMate);

                        SolidWorksDocumentProvider.GetActiveDoc().EditRebuild3();
                    }


                    // 3. Add compartmentConfig
                    compartmentsManager.ActivateDocument();

                    compartmentsManager.AddCompartment(
                        _projectFolder,
                        _compartmentDocPath,
                        compartmentDishedEndReference,
                        addedCompartmentConfig.config.Length / 1000);

                    // Reposition the compartmentConfig in SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments list
                    Compartment newCompartment = compartments.Last();
                    newCompartment._compartmentSettings.ID = addedCompartmentConfig.config.ID;
                    compartments.RemoveAt(compartments.Count - 1);
                    compartments.Insert(addedCompartmentConfig.index, newCompartment);

                    CompartmentsMappingHelper.AddOrUpdateMapping(newCompartment, addedCompartmentConfig.config);
                    CompartmentsMappingHelper.SaveOriginalConfigurations();

                    compartmentsManager.CloseDocument();

                    ModelDoc2 topLevelModel = SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc;
                    topLevelModel.EditRebuild3();
                    topLevelModel.ForceRebuild3(true);
                    topLevelModel.GraphicsRedraw2();
                }
            }

            if (switchedConfigs.Count > 0)
            {
                _hasChanges = true;
                // Reorder list of switched configs according indeces from lowest to highest
                switchedConfigs = switchedConfigs.OrderBy(item => item.index).ToList();

                foreach ((CompartmentConfiguration config, int index) switchedConfig in switchedConfigs)
                {
                    // Find the compartmentConfig to switch
                    (Compartment compartment, CompartmentConfiguration config)? mapping =
                        CompartmentsMappingHelper.GetMappingById(switchedConfig.config.ID);

                    int compartmentIndexToSwapWith = 0;
                    int compartmentToSwapCurrentIndex = 0;
                    // Find compartmentConfig's current index
                    for (int i = 0; i < SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments.Count; i++)
                    {
                        if (SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[i]._compartmentSettings.ID ==
                            mapping.Value.compartment._compartmentSettings.ID)
                        {
                            compartmentToSwapCurrentIndex = i;
                            compartmentIndexToSwapWith = switchedConfig.index;

                            // Continue if the compartmentConfig has been already swapped
                            if (compartmentToSwapCurrentIndex == compartmentIndexToSwapWith)
                            {
                                compartmentIndexToSwapWith = -1;
                                break;
                            }
                            break;
                        }

                    }

                    if (compartmentIndexToSwapWith >= 0)
                    {
                        SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.SwapCompartments(
                        compartmentToSwapCurrentIndex,
                        compartmentIndexToSwapWith);


                        SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.ActivateDocument();

                        if (compartmentToSwapCurrentIndex == SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments.Count - 1)
                            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(
                                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Last().Length / 1000);
                        else
                            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[compartmentToSwapCurrentIndex].ChangeDistance(
                                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[compartmentToSwapCurrentIndex].Length / 1000);

                        if (switchedConfig.index == SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments.Count - 1)
                            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(
                                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Last().Length / 1000);
                        else
                            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[switchedConfig.index].ChangeDistance(
                                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[switchedConfig.index].Length / 1000);

                        DocumentManager.UpdateAndSaveDocuments();
                    }
                }

                CompartmentsMappingHelper.ClearMappings();
                for (int i = 0; i < SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments.Count; i++)
                {
                    Compartment compartment = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[i];
                    CompartmentConfiguration compartmentConfiguration = SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i];

                    CompartmentsMappingHelper.AddOrUpdateMapping(compartment, compartmentConfiguration);
                }

                CompartmentsMappingHelper.SaveOriginalConfigurations();
            }
            if (_hasChanges == false)
            {
                SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.CloseDocument(); ;
                return;
            }
            

            ChangeCylindricalShellLength();

            //Save original configurations in mapping, so further modifications could be traced
            CompartmentsMappingHelper.SaveOriginalConfigurations();

            DocumentManager.DeleteFiles();

            DocumentManager.UpdateAndSaveDocuments();
        }

        /// <summary>
        /// Updates compartments configurations based on detected modifications.
        /// Identifies compartments with changes in length or left dished end alignment,
        /// and applies the necessary updates to synchronize their states with their configurations.
        /// Finally, saves the updated configurations for future tracking.
        /// </summary>
        public void ApplyCompartmentsChanges()
        {
            // Get modified configuration IDs
            List<Guid> modifiedConfigurationIds = CompartmentsMappingHelper.GetModifiedConfigurations();
            if (modifiedConfigurationIds.Count == 0) return;

            // Prepare lists for modifications
            List<(Compartment compartment, CompartmentConfiguration config)> lengthChanges = new List<(Compartment compartment, CompartmentConfiguration config)>();
            List<(Compartment compartment, CompartmentConfiguration config)> alignmentChanges = new List<(Compartment compartment, CompartmentConfiguration config)>();

            // Process each modified configuration
            foreach (Guid configId in modifiedConfigurationIds)
            {
                (Compartment compartment, CompartmentConfiguration config)? mapping = CompartmentsMappingHelper.GetMappingById(configId);

                if (mapping.HasValue)
                {
                    var (compartment, config) = mapping.Value;

                    // Check for length differences
                    if (compartment.GetLength() != config.Length / 1000)
                        lengthChanges.Add((compartment, config));

                    // Check for alignment changes
                    if (LeftDishedEndAlignmentChanged(compartment, config))
                        alignmentChanges.Add((compartment, config));
                }
            }

            if (lengthChanges.Count == 0 && alignmentChanges.Count == 0)
                return;

            _hasChanges = true;

            // Apply length and alignment changes
            if (lengthChanges.Count > 0)
            {
                ChangeCompartmentsLengths(lengthChanges);
                ChangeDishedEndDistance(lengthChanges);
            }

            if (alignmentChanges.Count > 0)
            {
                ChangeLeftDishedEndsAlignments(alignmentChanges);
            }
        }

        /// <summary>
        /// Updates the names of all compartments in the `CompartmentConfigurations` list.
        /// Names are sequentially assigned as "Compartment A", "Compartment B", and so on, 
        /// based on their index in the list. If there are more than 26 compartments, names are "Compartment 1", "Compartment 2", etc.
        /// Calls `TankSiteDataManager.UpdateTankProperties` to apply changes to tank properties.
        /// </summary>
        public void AdjustCompartmentsNames()
        {
            if (CompartmentConfigurations == null || CompartmentConfigurations.Count == 0)
            {
                return;
            }

            //Cannot assign compartmentConfig names beyond 'Z'.
            if (CompartmentConfigurations.Count <= 26)
            {
                for (int i = 0; i < CompartmentConfigurations.Count; i++)
                {
                    CompartmentConfigurations[i].Name = $"Compartment {(char)('A' + i)}";
                }
            }
            else
            {
                for (int i = 0; i < CompartmentConfigurations.Count; i++)
                {
                    CompartmentConfigurations[i].Name = $"Compartment {i}";
                }
            }


            try
            {
                TankSiteDataManager.UpdateTankProperties();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving the document: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a new compartmentConfig configuration with default values, assigns it a unique ID, 
        /// and adds it to the compartments list in both `SolidWorksDocumentProvider` and the local collection.
        /// </summary>
        /// <returns>
        /// The newly created <see cref="CompartmentConfiguration"/>.
        /// </returns>
        public CompartmentConfiguration CreateCompartmentConfiguration()
        {
            // Validate dependencies
            if (SolidWorksDocumentProvider._tankProperties == null ||
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations == null)
            {
                MessageBox.Show($"Compartment configuration could not be created. Tank Properties or Compartments Configurations object is null.");
                return null;
            }

            // Create new compartmentConfig with default values
            CompartmentConfiguration compartmentConfiguration = new CompartmentConfiguration
            {
                ID = Guid.NewGuid(),
                Amount = 100,
                AmountUnits = PaintingAmountUnit.Percentage,
                Name = GetCompartmentName()
            };

            // Add to SolidWorks tank properties and local list
            SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Add(compartmentConfiguration);
            CompartmentConfigurations.Add(compartmentConfiguration);

            return compartmentConfiguration;
        }

        /// <summary>
        /// Refreshes the `CompartmentConfigurations` list by clearing it and repopulating it with 
        /// compartments from `SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations`.
        /// Ensures each compartmentConfig has its `Amount` set to 100 and `AmountUnits` set to `Percentage`.
        /// </summary>
        public void GetCompartmentConfigurations()
        {
            // Validate dependencies
            if (SolidWorksDocumentProvider._tankProperties == null ||
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations == null)
            {
                return;
            }

            // Clear the existing configurations
            CompartmentConfigurations.Clear();

            // Retrieve and update compartments
            foreach (CompartmentConfiguration compartmentConfig in SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations)
            {
                compartmentConfig.Amount = 100; // Default value
                compartmentConfig.AmountUnits = PaintingAmountUnit.Percentage; // Default units
                CompartmentConfigurations.Add(compartmentConfig);
            }
        }

        /// <summary>
        /// Clears the `InternalTreatments` list and repopulates it with treatments 
        /// retrieved from the `SettingsDataManager`. Only treatments of type 
        /// `Internal` or `None` are added.
        /// </summary>
        public void GetInternalTreatments()
        {
            // Clear the existing treatments
            InternalTreatments.Clear();

            // Retrieve treatments from the settings manager
            List<Treatment> availableTreatments = SettingsDataManager.GetTreatments();
            if (availableTreatments == null)
            {
                return;
            }

            // Filter and add treatments of type Internal or None
            foreach (Treatment treatment in availableTreatments)
            {
                if (treatment.Type == TreatmentType.Internal || treatment.Type == TreatmentType.None)
                    InternalTreatments.Add(treatment);
            }
        }

        public void CalculateVolume(CompartmentConfiguration compartmentConfiguration)
        {
            double newLength = compartmentConfiguration.Length / 1000;

            //Recalculate volume
            double cylindricalShellVolume = _cylindricalShellVolumePerMeter * newLength;

            // Check if there are more than 1 compartmentConfig
            if (CompartmentConfigurations.Count == 1)
            {
                compartmentConfiguration.Volume = Math.Round(cylindricalShellVolume + (_mainDishedEndIncludeVolume * 2), 2);
            }
            else
            {
                int compartmentsCount = CompartmentConfigurations.Count;
                for (int i = 0; i < compartmentsCount; i++)
                {
                    // Find the count number of the compartmentConfig
                    if (CompartmentConfigurations[i] == compartmentConfiguration)
                    {
                        // Check if the compartmentConfig is the first one
                        if (i == 0)
                        {
                            // Get inner dished ends volume
                            double innerDishedEndVolume =
                                CompartmentConfigurations[i + 1].LeftDishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            compartmentConfiguration.Volume = Math.Round(cylindricalShellVolume + (_mainDishedEndIncludeVolume + innerDishedEndVolume), 2);
                        }
                        // Check if the compartmentConfig is the last one
                        else if (i + 1 == compartmentsCount)
                        {
                            // Get inner dished ends volume
                            double innerDishedEndVolume =
                                CompartmentConfigurations[i].LeftDishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            compartmentConfiguration.Volume =
                                Math.Round(cylindricalShellVolume + (_mainDishedEndIncludeVolume + innerDishedEndVolume), 2);
                        }
                        else
                        {
                            double innerLeftDishedEndVolume =
                                CompartmentConfigurations[i].LeftDishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
                            double innerRightDishedEndVolume =
                                CompartmentConfigurations[i + 1].LeftDishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            compartmentConfiguration.Volume =
                                Math.Round(cylindricalShellVolume + (innerLeftDishedEndVolume + innerRightDishedEndVolume), 2);
                        }
                        break;
                    }

                }
            }
        }

        public void CalculateLength(CompartmentConfiguration compartmentConfiguration)
        {
            double requiredVolumeOfCylindrcialShell = 0;

            double newVolume = compartmentConfiguration.Volume;

            //Recalculate length
            // Get volume of cylindrical shell needed
            if (CompartmentConfigurations.Count == 1)
            {
                requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume * 2);
            }
            else
            {
                int compartmentsCount = CompartmentConfigurations.Count;

                for (int i = 0; i < compartmentsCount; i++)
                {
                    // Find the count number of the compartmentConfig compartmentConfig
                    if (CompartmentConfigurations[i] == compartmentConfiguration)
                    {
                        // Check if the compartmentConfig is the first one
                        if (i == 0)
                        {
                            // Get inner dished ends volume
                            double innerDishedEndVolume =
                                CompartmentConfigurations[i + 1].LeftDishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                        }
                        // Check if the compartmentConfig is the last one
                        else if (i + 1 == compartmentsCount)
                        {
                            // Get inner dished ends volume
                            double innerDishedEndVolume =
                                CompartmentConfigurations[i].LeftDishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                        }
                        else
                        {
                            double innerLeftDishedEndVolume =
                                CompartmentConfigurations[i].LeftDishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
                            double innerRightDishedEndVolume =
                                CompartmentConfigurations[i + 1].LeftDishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                            requiredVolumeOfCylindrcialShell = newVolume - (innerLeftDishedEndVolume + innerRightDishedEndVolume);
                        }
                        break;
                    }
                }
            }

            // Get length of cylindrical shell
            compartmentConfiguration.Length = (int)(requiredVolumeOfCylindrcialShell / _cylindricalShellVolumePerMeter * 1000);
        }

        /// <summary>
        /// Removes the specified compartment configuration from the system, updates mappings, 
        /// reassigns names, and adjusts properties for the remaining compartments.
        /// </summary>
        /// <param name="compartmentConfiguration">
        /// The compartment configuration to be removed.
        /// </param>
        public void RemoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration)
        {
            if (compartmentConfiguration == null)
            {
                return;
            }

            if (!CompartmentConfigurations.Contains(compartmentConfiguration))
            {
                return;
            }

            int compartmentConfigIndex = CompartmentConfigurations.IndexOf(compartmentConfiguration);

            // Remove mappings and the compartment from the lists
            CompartmentsMappingHelper.RemoveMapping(compartmentConfiguration.ID);
            CompartmentConfigurations.Remove(compartmentConfiguration);
            SolidWorksDocumentProvider._tankProperties?.CompartmentsConfigurations.Remove(compartmentConfiguration);

            // Adjust the first compartment's properties if the list is not empty
            if (CompartmentConfigurations.Count > 0)
            {
                CompartmentConfigurations[0].LeftDishedEndAlignment = DishedEndAlignment.Left;
                CompartmentConfigurations[0].LeftEndConnection = LeftEndConnection.None;
            }

            if (CompartmentConfigurations.Count == 0) return;

            if (compartmentConfigIndex == 0) CalculateVolume(CompartmentConfigurations[compartmentConfigIndex]);

            else if (compartmentConfigIndex == CompartmentConfigurations.Count) CalculateVolume(CompartmentConfigurations[compartmentConfigIndex - 1]);

            else
            {
                CalculateVolume(CompartmentConfigurations[compartmentConfigIndex - 1]);
                CalculateVolume(CompartmentConfigurations[compartmentConfigIndex]);
            }

            // Reassign compartment names
            AdjustCompartmentsNames();
        }

        /// <summary>
        /// Updates the `LeftEndConnection` property of the specified compartment configuration 
        /// with a new connection type.
        /// </summary>
        /// <param name="compartmentConfiguration">
        /// The target <see cref="CompartmentConfiguration"/> to update.
        /// </param>
        /// <param name="newConnection">
        /// The new <see cref="LeftEndConnection"/> value to assign.
        /// </param>
        public void UpdateLeftEndConnection(CompartmentConfiguration compartmentConfiguration, LeftEndConnection newConnection)
        {
            if (compartmentConfiguration == null)
            {
                return;
            }

            if (!CompartmentConfigurations.Contains(compartmentConfiguration))
            {
                return;
            }

            // Update the connection
            compartmentConfiguration.LeftEndConnection = newConnection;
        }

        /// <summary>
        /// Updates the `LeftDishedEndAlignment` property of the specified compartment configuration 
        /// with a new alignment value.
        /// </summary>
        /// <param name="compartmentConfiguration">
        /// The target <see cref="CompartmentConfiguration"/> to update.
        /// </param>
        /// <param name="newAlignment">
        /// The new <see cref="DishedEndAlignment"/> value to assign.
        /// </param>
        public void ChangeLeftDishedEndAlignment(CompartmentConfiguration compartmentConfiguration, DishedEndAlignment newAlignment)
        {
            if (compartmentConfiguration == null) return;

            if (!CompartmentConfigurations.Contains(compartmentConfiguration)) return;

            // Update the alignment
            compartmentConfiguration.LeftDishedEndAlignment = newAlignment;

            int compartmentConfigIndex = CompartmentConfigurations.IndexOf(compartmentConfiguration);

            if (compartmentConfigIndex > 0)
            {
                CalculateVolume(CompartmentConfigurations[compartmentConfigIndex - 1]);
                CalculateVolume(compartmentConfiguration);
            }
        }

        /// <summary>
        /// Moves a specified compartment configuration to a new index within the list of 
        /// compartment configurations in both `CompartmentConfigurations` and 
        /// `SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations`.
        /// </summary>
        /// <param name="compartmentConfiguration">
        /// The target <see cref="CompartmentConfiguration"/> to move.
        /// </param>
        /// <param name="newIndex">
        /// The new index position where the compartment configuration should be inserted.
        /// </param>
        public void MoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration, int newIndex)
        {
            if (compartmentConfiguration == null) return;
            if (!CompartmentConfigurations.Contains(compartmentConfiguration)) return;
            if (newIndex < 0 || newIndex >= CompartmentConfigurations.Count) return;

            // Remove from both lists
            CompartmentConfigurations.Remove(compartmentConfiguration);
            SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Remove(compartmentConfiguration);

            // Insert into the new position in both lists
            CompartmentConfigurations.Insert(newIndex, compartmentConfiguration);

            if (newIndex > 0) CalculateVolume(CompartmentConfigurations[newIndex - 1]);
            CalculateVolume(CompartmentConfigurations[newIndex]);

            SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Insert(newIndex, compartmentConfiguration);
        }

        public bool HasChanges()
        {
            // Anything changed in the configuration list?
            if (CompartmentsMappingHelper.GetDeletedConfigurations().Count > 0) return true;
            if (CompartmentsMappingHelper.GetAddedConfigurations().Count > 0) return true;
            if (CompartmentsMappingHelper.GetSwitchedConfigurations(CompartmentConfigurations).Count > 0) return true;

            // Any edits inside existing configs (length/alignment/etc.)?
            if (CompartmentsMappingHelper.GetModifiedConfigurations().Count > 0) return true;

            return false;
        }

        public void BackButtonClick()
        {
        }
    }

}

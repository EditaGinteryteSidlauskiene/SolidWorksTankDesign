using AddinWithTaskpane;
using Microsoft.VisualBasic.FileIO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    public class Compartment
    {
        private ModelDoc2 _currentlyActiveCompartmentDoc;

        private const string COMPARTMENT_COMPONENT_NAME = "Compartment";
        private const string LEFT_END_PLANE_NAME = "Dished end position plane";
        private const string CENTER_AXIS_NAME = "Center axis";
        private const string FRONT_PLANE_NAME = "Front plane";

        public List<Nozzle> Nozzles = new List<Nozzle>();

        public CompartmentSettings _compartmentSettings;

        public Compartment() 
        {
            _compartmentSettings = new CompartmentSettings();
        }

        /// <summary>
        /// This constructor is called when a new compartment is added.
        /// </summary>
        /// <param name="shellFrontPlane"></param>
        /// <param name="shellCenterAxis"></param>
        /// <param name="dishedEndPositionPlane"></param>
        /// <param name="countNumber"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public Compartment(
            string projectFolder,
            string compartmentPath,
            Feature shellFrontPlane,
            Feature shellCenterAxis,
            Feature dishedEndPositionPlane,
            double length)
        {
            // Input Validation
            if (shellCenterAxis == null)
                throw new ArgumentNullException(nameof(shellCenterAxis));

            if (shellFrontPlane == null)
                throw new ArgumentNullException(nameof(shellFrontPlane));

            if (dishedEndPositionPlane == null)
                throw new ArgumentNullException(nameof(dishedEndPositionPlane));

            // Activate shell doc
            ModelDoc2 shellModelDoc = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();
            if (shellModelDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            // Rename the Component
            string componentName = COMPARTMENT_COMPONENT_NAME;

            // Open compartmen doc
            ModelDoc2 compartmentModelDoc2 = SolidWorksDocumentProvider._solidWorksApplication.OpenDoc6(
               compartmentPath,
               (int)swDocumentTypes_e.swDocASSEMBLY,
               (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
               "", 1, 1);

            // Package the nozzle assembly and its associated files using Pack and Go, and get the path to the packed assembly
            string path = DocumentManager.PackAndGo(projectFolder, compartmentModelDoc2, componentName, null);

            string docName = path.Split('\\').Last();
            // Rename the doc
            string filePath = Path.GetDirectoryName(path) + "\\Compartment.SLDASM";
            FileSystem.RenameFile(filePath, docName);
            // Construct the new file path
            //string newPath = Path.Combine(Path.GetDirectoryName(path), $"{componentName}.SLDASM");

            // Close the nozzle assembly document after it has been packed
            string compartmentModelDoc2Title = compartmentModelDoc2.GetTitle();
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(compartmentModelDoc2Title);
            Marshal.ReleaseComObject(compartmentModelDoc2);
            compartmentModelDoc2 = null;

            // 2. Add and Make Independent Compartment Component
            Component2 compartment = ComponentManager.AddComponentAssembly(shellModelDoc, path);

            ModelDoc2 compartmentModelDoc = compartment.GetModelDoc2();

            Feature leftEndMate = null;
            Feature frontPlaneMate = null;
            Feature centerAxisMate = null;

            // 6. Create Mates
            try
            {
                leftEndMate = MateManager.CreateMate(
                    componentFeature1: dishedEndPositionPlane,
                    componentFeature2: SWFeatureManager.GetMajorPlane(compartment, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{compartment.Name2} - {LEFT_END_PLANE_NAME}");

                centerAxisMate = MateManager.CreateMate(
                    componentFeature1: shellCenterAxis,
                    componentFeature2: SWFeatureManager.GetFeatureByName(compartment, "Center axis"),
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{compartment.Name2} - {CENTER_AXIS_NAME}");


                frontPlaneMate = MateManager.CreateMate(
                    componentFeature1: shellFrontPlane,
                    componentFeature2: SWFeatureManager.GetMajorPlane(compartment, MajorPlane.Front),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{compartment.Name2} - {FRONT_PLANE_NAME}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // 7. Get compartment Entities and Initialize Settings
            _compartmentSettings = new CompartmentSettings();
            try
            {
                GetCompartmentPIDsAndChangeLength();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting compartment entities: {ex.Message}");
            }

            void GetCompartmentPIDsAndChangeLength()
            {
                try
                {
                    // Populate the _compartmentSettings with the retrieved PIDs
                    _compartmentSettings.PIDComponent = shellModelDoc.Extension.GetPersistReference3(compartment);
                    _compartmentSettings.PIDLeftEndMate = shellModelDoc.Extension.GetPersistReference3(leftEndMate);
                    _compartmentSettings.PIDFrontPlaneMate = shellModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                    _compartmentSettings.PIDCenterAxisMate = shellModelDoc.Extension.GetPersistReference3(centerAxisMate);
                    _compartmentSettings.PIDDishedEndPositionPlane = shellModelDoc.Extension.GetPersistReference3(dishedEndPositionPlane);

                    using (var compartmentDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, compartmentModelDoc))
                    {
                        Feature leftEndPlane = SWFeatureManager.GetFeatureByName(compartmentModelDoc, "Left end plane");
                        Feature rightEndPlane = SWFeatureManager.GetFeatureByName(compartmentModelDoc, "Right end plane");
                        Feature compartmentCenterAxis = SWFeatureManager.GetFeatureByName(compartmentModelDoc, "Center axis");

                        _compartmentSettings.PIDCenterAxis = compartmentModelDoc.Extension.GetPersistReference3(compartmentCenterAxis);
                        _compartmentSettings.PIDLeftEndPlane = compartmentModelDoc.Extension.GetPersistReference3(leftEndPlane);
                        _compartmentSettings.PIDRightEndPlane = compartmentModelDoc.Extension.GetPersistReference3(rightEndPlane);

                        // Change compartment's length
                        ChangeLength(length);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Could not initialize compartment settings.");
                    return;
                }
            }

            DocumentManager.UpdateAndSaveDocuments();
        }

        public void ChangeLength(double length)
        {
            Feature rightEndPlane = GetRightEndPlane();

            SWFeatureManager.ChangeDistanceOfReferencePlane(rightEndPlane, length);

            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();

            compartmentDoc.Save3(
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                            (int)swFileSaveError_e.swGenericSaveError,
                            (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        /// <summary>
        /// Deletes a compartment component from a SolidWorks shell assembly, including the associated file and folder that contains it.
        /// </summary>
        public void Delete()
        {
            ModelDoc2 shellModelDoc = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();

            shellModelDoc.ClearSelection2(true);

            SelectionMgr selectionManager = (SelectionMgr)shellModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            // Get compartment's cut out planes
            if(Nozzles.Count > 0)
            {
                foreach(Nozzle nozzle in Nozzles)
                {
                    nozzle.GetCutOutPlane().Select2(true, 1);
                }

                shellModelDoc.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Children);
            }

            //Select the compartment to be deleted
            GetComponent().Select4(false, selectData, false);

            //Get compartment document's path to delete the file
            ModelDoc2 componentDocument = GetComponent().GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected compartment
            ((AssemblyDoc)shellModelDoc).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            shellModelDoc.EditRebuild3();

            DocumentManager.UpdateAndSaveDocuments();

            //Delete the file
            File.Delete(path);

            // Get the parent folder and delete it
            string folderPath = Path.GetDirectoryName(path);
            Directory.Delete(folderPath);
        }

        public double GetLength()
        {
            ActivateDocument();

            double length = SWFeatureManager.GetDistanceOfReferencePlane(GetRightEndPlane());

            CloseDocument();

            return length;
        }

        /// <summary>
        /// Adds nozzle in the current compartment's document and adds it to the list.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        /// <param name="referencePlane"></param>
        /// <param name="distance"></param>
        public Nozzle AddNozzle(
            string nozzlePositionSketchPath,
            string nozzleDocPath,
            int compartmentNumber,
            Feature referencePlane,
            double distance,
            bool flip,
            double externalDiameter)
        {
            Nozzle nozzle = null;
            try
            {
                nozzle = new Nozzle(
                        nozzlePositionSketchPath,
                        compartmentNumber,
                        referencePlane,
                        GetCenterAxis(),
                        SWFeatureManager.GetMajorPlane(_currentlyActiveCompartmentDoc, MajorPlane.Front),
                        distance,
                        flip,
                        externalDiameter);

                Nozzles.Add(nozzle);

                Compartment compartment = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber];

                nozzle.AddNozzleAssembly(nozzleDocPath, compartment, nozzle);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            DocumentManager.UpdateAndSaveDocuments();

            return nozzle;
        }

        /// <summary>
        /// Removes the last nozzle component from the compartment assembly.
        /// </summary>
        public bool DeleteNozzle()
        {
            // Get the count of nozzles in the current compartment once for efficiency
            int nozzlesCount = Nozzles.Count;

            // Check if there are enough nozzles to remove
            if (nozzlesCount == 0)
            {
                MessageBox.Show($"There are no nozzles in the compartment assembly.");
                CloseDocument();
                return false;
            }

            try
            {
                SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();
                Nozzles[nozzlesCount - 1].DeleteCutExtrude();

                // Activates compartment document
                ActivateDocument();
                // Attempt to delete the SolidWorks object associated with the compartment
                Nozzles[nozzlesCount - 1].DeleteNozzle();
            }
            catch (Exception ex)
            {
                // Handle potential exceptions during deletion
                MessageBox.Show($"Error deleting nozzle: {ex.Message}");
                return false;
            }

            // Remove the nozzle from the internal tracking list
            Nozzles.RemoveAt(nozzlesCount - 1);

            // Update documents and close assembly of compartment
            DocumentManager.UpdateAndSaveDocuments();
            _currentlyActiveCompartmentDoc = null;

            return true;
        }

        /// <summary>
        /// Activates document of compartment assembly
        /// </summary>
        public ModelDoc2 ActivateDocument()
        {
            // Activate shell doc
            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();

            // Get compartment doc
            ModelDoc2 compartmentModelDoc = GetComponent().GetModelDoc2();

            // Activate compartment doc
            _currentlyActiveCompartmentDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(compartmentModelDoc.GetTitle() + ".sldasm", true, 0, 0);
        
            return compartmentModelDoc;
        }

        /// <summary>
        /// Closes active document of compartment assembly
        /// </summary>
        public void CloseDocument()
        {
            if (_currentlyActiveCompartmentDoc == null) return;

            _currentlyActiveCompartmentDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(_currentlyActiveCompartmentDoc.GetTitle());
            _currentlyActiveCompartmentDoc = null;
        }

        public Feature GetLeftEndPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                       _compartmentSettings.PIDLeftEndPlane,
                       out int error);

        public Feature GetRightEndPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDRightEndPlane,
                        out int error);

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDComponent,
                        out int error);

        public Feature GetCenterAxis() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDCenterAxis,
                        out int error);

        public Feature GetLeftEndMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDLeftEndMate,
                        out int error);

        public Feature GetFrontPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDFrontPlaneMate,
                        out int error);

        public Feature GetCenterAxisMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDCenterAxisMate,
                        out int error);

        public Feature GetDishedEndPositionPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDDishedEndPositionPlane,
                        out int error);
    }
}

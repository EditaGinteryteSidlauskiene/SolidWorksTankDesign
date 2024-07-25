using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class Compartment
    {
        private ModelDoc2 _currentlyActiveCompartmentDoc;

        private const string COMPARTMENT_ASSEMBLY_PATH = "C:\\Users\\Edita\\TankDesignStudio\\ClassA\\d2500\\Shell\\Empty Compartment.SLDASM";
        private const string COMPARTMENT_COMPONENT_NAME = "Compartment";
        private const string LEFT_END_PLANE_NAME = "Left end plane";
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
            Feature shellFrontPlane,
            Feature shellCenterAxis,
            Feature dishedEndPositionPlane,
            int countNumber,
            double distanceBetweenNozzleAndRefPlane)
        {
            // 1. Input Validation
            if (shellCenterAxis == null)
                throw new ArgumentNullException(nameof(shellCenterAxis));

            if (shellFrontPlane == null)
                throw new ArgumentNullException(nameof(shellFrontPlane));

            if (dishedEndPositionPlane == null)
                throw new ArgumentNullException(nameof(dishedEndPositionPlane));

            ModelDoc2 shellModelDoc = SolidWorksDocumentProvider.GetActiveDoc();
            if (shellModelDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            // 2. Add and Make Independent Compartment Component
            Component2 compartment = ComponentManager.AddComponentAssembly(shellModelDoc, COMPARTMENT_ASSEMBLY_PATH);
            ComponentManager.MakeComponentIndependent(compartment, COMPARTMENT_ASSEMBLY_PATH);

            // 3. Rename the Component
            char letter = (char)(65 + countNumber);
            string componentName = $"{COMPARTMENT_COMPONENT_NAME} {letter}";
            FeatureManager.GetFeatureByName(shellModelDoc, compartment.Name2).Name = componentName;

            ModelDoc2 compartmentModelDoc = compartment.GetModelDoc2();

            Feature leftEndMate = null;

            // 6. Create Mates
            try
            {
                leftEndMate = MateManager.CreateMate(
                    componentFeature1: dishedEndPositionPlane,
                    componentFeature2: FeatureManager.GetMajorPlane(compartment, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{compartment.Name2} - {LEFT_END_PLANE_NAME}");

                MateManager.CreateMate(
                    componentFeature1: shellCenterAxis,
                    componentFeature2: FeatureManager.GetFeatureByName(compartment, "Center axis"),
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{compartment.Name2} - {CENTER_AXIS_NAME}");


                MateManager.CreateMate(
                    componentFeature1: shellFrontPlane,
                    componentFeature2: FeatureManager.GetMajorPlane(compartment, MajorPlane.Front),
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
                GetCompartmentPIDs();

                ActivateDocument();

                AddNozzle(
                    countNumber,
                    GetRightEndPlane(),
                    distanceBetweenNozzleAndRefPlane);

                CloseDocument();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting compartment entities: {ex.Message}");
            }

            void GetCompartmentPIDs()
            {
                try
                {
                    // Populate the _compartmentSettings with the retrieved PIDs
                    _compartmentSettings.PIDComponent = shellModelDoc.Extension.GetPersistReference3(compartment);
                    _compartmentSettings.PIDLeftEndMate = shellModelDoc.Extension.GetPersistReference3(leftEndMate);

                    using (var compartmentDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, compartmentModelDoc))
                    {
                        Feature leftEndPlane = FeatureManager.GetFeatureByName(compartmentModelDoc, "Left end plane");
                        Feature rightEndPlane = FeatureManager.GetFeatureByName(compartmentModelDoc, "Right end plane");
                        Feature compartmentCenterAxis = FeatureManager.GetFeatureByName(compartmentModelDoc, "Center axis");

                        _compartmentSettings.PIDCenterAxis = compartmentModelDoc.Extension.GetPersistReference3(compartmentCenterAxis);
                        _compartmentSettings.PIDLeftEndPlane = compartmentModelDoc.Extension.GetPersistReference3(leftEndPlane);
                        _compartmentSettings.PIDRightEndPlane = compartmentModelDoc.Extension.GetPersistReference3(rightEndPlane);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Could not initialize compartment settings.");
                    return;
                }
            }
        }

        /// <summary>
        /// Deletes a compartment component from a SolidWorks shell assembly, including the associated file.
        /// </summary>
        public void Delete()
        {
            ModelDoc2 shellModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr selectionManager = (SelectionMgr)shellModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            GetComponent().Select4(false, selectData, false);

            //Get compartment document's path to delete the file
            ModelDoc2 componentDocument = GetComponent().GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected compartment
            ((AssemblyDoc)shellModelDoc).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            shellModelDoc.EditRebuild3();

            //Delete the file
            File.Delete(path);
        }

        /// <summary>
        /// Adds nozzle in the current compartment's document and adds it to the list.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        /// <param name="referencePlane"></param>
        /// <param name="distance"></param>
        public void AddNozzle(
            int compartmentNumber,
            Feature referencePlane,
            double distance)
        {
            try
            {
                Nozzles.Add(
                    new Nozzle(
                        compartmentNumber,
                        referencePlane,
                        GetCenterAxis(),
                        FeatureManager.GetMajorPlane(_currentlyActiveCompartmentDoc, MajorPlane.Front),
                        distance));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Activates document of compartment assembly
        /// </summary>
        public void ActivateDocument()
        {
            // Activate shell doc
            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();

            // Get compartment doc
            ModelDoc2 compartmentModelDoc = GetComponent().GetModelDoc2();

            // Close shell doc
            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.CloseDocument();

            // Activate compartment doc
            _currentlyActiveCompartmentDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(compartmentModelDoc.GetTitle() + ".sldasm", true, 0, 0);
        }

        /// <summary>
        /// Closes active document of compartment assembly
        /// </summary>
        public void CloseDocument()
        {
            if (_currentlyActiveCompartmentDoc == null) return;

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
    }
}

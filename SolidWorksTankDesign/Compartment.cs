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
        private const string COMPARTMENT_ASSEMBLY_PATH = "C:\\Users\\Edita\\TankDesignStudio\\ClassA\\d2500\\Shell\\Compartment A Manholes.SLDASM";
        private const string COMPARTMENT_COMPONENT_NAME = "Compartment";
        private const string LEFT_END_PLANE_NAME = "Left end plane";
        private const string CENTER_AXIS_NAME = "Center axis";
        private const string FRONT_PLANE_NAME = "Front plane";

        public List<object> Manholes = new List<object>();

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
            int countNumber)
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

            // 4. Get Features for Mating
            Feature leftEndPlane = FeatureManager.GetFeatureByName(compartment, "Left end plane");
            Feature rightEndPlane = FeatureManager.GetFeatureByName(compartment, "Right end plane");
            Feature compartmentCenterAxis = FeatureManager.GetFeatureByName(compartment, "Center axis");

            Feature leftEndMate = null;
            Feature centerAxisMate = null;
            Feature frontPlaneMate = null;

            // 6. Create Mates
            try
            {
                leftEndMate = MateManager.CreateMate(
                componentFeature1: dishedEndPositionPlane,
                componentFeature2: leftEndPlane,
                alignmentType: MateAlignment.Aligned,
                name: $"{compartment.Name2} - {LEFT_END_PLANE_NAME}");

                centerAxisMate = MateManager.CreateMate(
                    componentFeature1: shellCenterAxis,
                    componentFeature2: compartmentCenterAxis,
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{compartment.Name2} - {CENTER_AXIS_NAME}");


                frontPlaneMate = MateManager.CreateMate(
                    componentFeature1: shellFrontPlane,
                    componentFeature2: FeatureManager.GetMajorPlane(compartment, MajorPlane.Front),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{compartment.Name2} - {FRONT_PLANE_NAME}"); ;
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
                    _compartmentSettings.PIDLeftEndPlane = shellModelDoc.Extension.GetPersistReference3(leftEndPlane);
                    _compartmentSettings.PIDRightEndPlane = shellModelDoc.Extension.GetPersistReference3(rightEndPlane);
                    _compartmentSettings.PIDCenterAxis = shellModelDoc.Extension.GetPersistReference3(compartmentCenterAxis);
                    _compartmentSettings.PIDComponent = shellModelDoc.Extension.GetPersistReference3(compartment);
                    _compartmentSettings.PIDCenterAxisMate = shellModelDoc.Extension.GetPersistReference3(centerAxisMate);
                    _compartmentSettings.PIDLeftEndMate = shellModelDoc.Extension.GetPersistReference3(leftEndMate);
                    _compartmentSettings.PIDFrontPlaneMate = shellModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "The attribute could not be created.");
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

        public Feature GetCenterAxisMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDCenterAxisMate,
                        out int error);

        public Feature GetLeftEndMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDLeftEndMate,
                        out int error);

        public Feature GetFrontPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _compartmentSettings.PIDFrontPlaneMate,
                        out int error);
    }
}

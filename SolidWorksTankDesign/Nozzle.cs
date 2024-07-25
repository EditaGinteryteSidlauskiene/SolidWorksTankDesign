using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class Nozzle
    {
        private ModelDoc2 _currentlyActiveNozzleDoc;

        private const string MANHOLE_NAME = "M";
        private const string POSITION_PLANE_NAME = "position plane";
        private const string CENTER_AXIS_NAME = "Center axis";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string TOP_PLANE_NAME = "Top plane";
        private const string NOZZLE_PATH = "C:\\Users\\Edita\\TankDesignStudio\\ClassA\\d2500\\Nozzles\\M1 Manhole.SLDASM";
        private const string NOZZLE_ASSEMBLY_PATH = "C:\\Users\\Edita\\Desktop\\Automation tank\\Manhole DN600 Neck with flange.SLDASM";

        public NozzleSettings _nozzleSettings;

        public Nozzle() 
        { 
            _nozzleSettings = new NozzleSettings();
        }

        /// <summary>
        /// This constructor is responsible for creating and positioning a nozzle within a SolidWorks assembly.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        /// <param name="referencePlane"></param>
        /// <param name="compartmentCenterAxis"></param>
        /// <param name="compartmentFrontPlane"></param>
        /// <param name="distance"></param>
        public Nozzle(
            int compartmentNumber,
            Feature referencePlane,
            Feature compartmentCenterAxis,
            Feature compartmentFrontPlane,
            double distance)
        {
            // Get active compartment's document
            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();

            string positionPlaneName = $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber-1].Nozzles.Count+1} {POSITION_PLANE_NAME}";
            // Create nozzle's position plane
            Feature positionPlane = FeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referencePlane,
                distance: distance,
                name: positionPlaneName);

            // Add a new nozzle and make it independent
            Component2 nozzle = ComponentManager.AddComponentAssembly(compartmentDoc, NOZZLE_PATH);
            ComponentManager.MakeComponentIndependent(nozzle, NOZZLE_PATH);

            // Rename nozzle component
            string componentName = $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber-1].Nozzles.Count+1}";
            FeatureManager.GetFeatureByName(compartmentDoc, nozzle.Name2).Name = componentName;

            ModelDoc2 nozzleModelDoc = nozzle.GetModelDoc2();

            MateNozzle();

            // Get nozzle Entities and Initialize Settings
            _nozzleSettings = new NozzleSettings();
            try
            {
                GetNozzlePIDs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting nozzle entities: {ex.Message}");
            }

            try
            {
                AddNozzleAssembly();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding nozzle assembly: {ex.Message}");
            }

            DocumentManager.UpdateAndSaveDocuments(compartmentDoc);

            void MateNozzle()
            {
                try
                {
                    MateManager.CreateMate(
                        componentFeature1: positionPlane,
                        componentFeature2: FeatureManager.GetMajorPlane(nozzle, MajorPlane.Right),
                        alignmentType: MateAlignment.Aligned,
                        name: $"{nozzle.Name2} - {POSITION_PLANE_NAME}");

                    MateManager.CreateMate(
                        componentFeature1: compartmentCenterAxis,
                        componentFeature2: FeatureManager.GetFeatureByName(nozzle, "Center Axis"),
                        alignmentType: MateAlignment.Anti_Aligned,
                        name: $"{nozzle.Name2} - {CENTER_AXIS_NAME}");

                    MateManager.CreateMate(
                        componentFeature1: compartmentFrontPlane,
                        componentFeature2: FeatureManager.GetMajorPlane(nozzle, MajorPlane.Front),
                        alignmentType: MateAlignment.Aligned,
                        name: $"{nozzle.Name2} - {FRONT_PLANE_NAME}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "At least one of nozzle mates could not be created.");
                    return;
                }
            }

            void GetNozzlePIDs()
            {
                try
                {
                    _nozzleSettings.PIDPositionPlane = compartmentDoc.Extension.GetPersistReference3(positionPlane);
                    _nozzleSettings.PIDComponent = compartmentDoc.Extension.GetPersistReference3(nozzle);

                    // Use a SolidWorksDocumentWrapper for managing the nozzle's model document.
                    using (var nozzleDocument = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, nozzleModelDoc))
                    {
                        // Get the selection manager to interact with selections within the nozzle's model
                        SelectionMgr selectionMgrAtNozzle = (SelectionMgr)nozzleModelDoc.SelectionManager;

                        // Get features and components
                        nozzleModelDoc.Extension.SelectByID2(
                            "Center Axis",
                            "AXIS",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature nozzleCenterAxis = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "External point",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature externalPoint = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "Internal point",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature internalPoint = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "Inside point",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature insidePoint = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "Mid point",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature midPoint = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "Nozzle Right Reference Plane",
                            "PLANE",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature nozzleRightRefPlane = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        //-------------- PERVADINTI --------------------------
                        nozzleModelDoc.Extension.SelectByID2(
                            "PLANE1",
                            "PLANE",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature plane1 = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        nozzleModelDoc.Extension.SelectByID2(
                            "Sketch",
                            "SKETCH",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature sketch = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        try
                        {
                            // Populate the _nozzleSettings with the retrieved PIDs for those entities that has to be reachable from nozzle's document
                            _nozzleSettings.PIDCenterAxis = nozzleModelDoc.Extension.GetPersistReference3(nozzleCenterAxis);
                            _nozzleSettings.PIDExternalPoint = nozzleModelDoc.Extension.GetPersistReference3(externalPoint);
                            _nozzleSettings.PIDInternalPoint = nozzleModelDoc.Extension.GetPersistReference3(internalPoint);
                            _nozzleSettings.PIDInsidePoint = nozzleModelDoc.Extension.GetPersistReference3(insidePoint);
                            _nozzleSettings.PIDMidPoint = nozzleModelDoc.Extension.GetPersistReference3(midPoint);
                            _nozzleSettings.PIDNozzleRightRefPlane = nozzleModelDoc.Extension.GetPersistReference3(nozzleRightRefPlane);
                            _nozzleSettings.PIDCutPlane = nozzleModelDoc.Extension.GetPersistReference3(plane1);
                            _nozzleSettings.PIDSketch = nozzleModelDoc.Extension.GetPersistReference3(sketch);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "The attribute could not be created.");
                            return;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Could not initialize nozzle settings.");
                    return;
                }
            }
        }

        /// <summary>
        /// This method inserts a pre-defined nozzle assembly into the currently active nozzle document. 
        /// It aligns and positions the assembly using mates, stores relevant persistent IDs, and handles potential errors during the process.
        /// Finally, it closes the active nozzle document.
        /// </summary>
        public void AddNozzleAssembly()
    {
        // Activate current nozzle document
            ActivateDocument();

            // Add nozzle assembly and make it independent
            Component2 nozzleAssembly = ComponentManager.AddComponentAssembly(_currentlyActiveNozzleDoc, NOZZLE_ASSEMBLY_PATH);
            ComponentManager.MakeComponentIndependent(nozzleAssembly, NOZZLE_ASSEMBLY_PATH);

            // Get features for mating
            Feature nozzleAssemblyCenterAxis = FeatureManager.GetFeatureByName(nozzleAssembly, "Center axis");

            // Create Mates
            try
            {
                MateManager.CreateMate(
                    componentFeature1: FeatureManager.GetFeatureByName(_currentlyActiveNozzleDoc, "Nozzle axis"),
                    componentFeature2: nozzleAssemblyCenterAxis,
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {CENTER_AXIS_NAME}");

                MateManager.CreateMate(
                    componentFeature1: GetNozzleRightRefPlane(),
                    componentFeature2: FeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {RIGHT_PLANE_NAME}");

                Feature topPlaneMate = MateManager.CreateMate(
                    componentFeature1: GetCutPlane(),
                    componentFeature2: FeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Top),
                    alignmentType: MateAlignment.Anti_Aligned,
                    distance: 0,
                    name: $"{nozzleAssembly.Name2} - {TOP_PLANE_NAME}");

                // Get nozzle assembly PIDs
                _nozzleSettings.PIDNozzleAssemblyComp = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(nozzleAssembly);
                _nozzleSettings.PIDTopPlaneMate = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "At least one of nozzle assembly mates could not be created.");
            }

            // Updates and saves attribute and nozzle document
            DocumentManager.UpdateAndSaveDocuments(_currentlyActiveNozzleDoc);
            CloseDocument();
        }

        /// <summary>
        /// Activates document of nozzle assembly
        /// </summary>
        public void ActivateDocument()
        {
            // Get compartment doc
            ModelDoc2 nozzleModelDoc = GetComponent().GetModelDoc2();

            // Activate compartment doc
            _currentlyActiveNozzleDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(nozzleModelDoc.GetTitle() + ".sldasm", true, 0, 0);
        }

        /// <summary>
        /// Closes active document of nozzle assembly
        /// </summary>
        public void CloseDocument()
        {
            if (_currentlyActiveNozzleDoc == null) return;

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(_currentlyActiveNozzleDoc.GetTitle());
            _currentlyActiveNozzleDoc = null;
        }

        public Feature GetCenterAxis() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                      _nozzleSettings.PIDCenterAxis,
                      out int error);

        public Feature GetPositionPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDPositionPlane,
                     out int error);

        public Feature GetExternalPoint() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDExternalPoint,
                     out int error);

        public Feature GetInternalPoint() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDInternalPoint,
                     out int error);

        public Feature GetInsidePoint() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDInsidePoint,
                     out int error);

        public Feature GetMidPoint() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDMidPoint,
                     out int error);

        public Feature GetNozzleRightRefPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDNozzleRightRefPlane,
                     out int error);

        //---------- PERVADINTI --------------------
        public Feature GetCutPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDCutPlane,
                     out int error);

        public Feature GetSketch() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDSketch,
                     out int error);

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDComponent,
                     out int error);

        public Feature GetTopPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDTopPlaneMate,
                     out int error);

        public Component2 GetNozzleAssemblyComp() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDNozzleAssemblyComp,
                     out int error);
    }
}

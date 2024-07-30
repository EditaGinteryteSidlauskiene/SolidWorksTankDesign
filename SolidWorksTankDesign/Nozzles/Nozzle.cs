using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Windows.Forms;
using WarningAndErrorService;

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
        private const string NOZZLE_PATH = "C:\\Users\\Edita\\TankDesignStudio\\ClassA\\d2500\\Nozzles\\Empty M1 Manhole.SLDASM";
        private const string NOZZLE_ASSEMBLY_PATH = "C:\\Users\\Edita\\Desktop\\Automation tank\\Manhole DN600 Neck with flange.SLDASM";
        private const string SHELL_DIAMETER_EXTERNAL = "ShellDiameterExternal";
        private const string SHELL_DIAMETER_INTERNAL = "ShellDiameterInternal";
        private const string CENTER_AXIS_ROTATION_ANGLE = "CenterAxisRotationAngle";
        private const string OFFSET = "Offset";

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

            Feature positionPlaneMate = null;
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

            void MateNozzle()
            {
                try
                {
                    positionPlaneMate = MateManager.CreateMate(
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
                    _nozzleSettings.PIDPositionPlaneMate = compartmentDoc.Extension.GetPersistReference3(positionPlaneMate);

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
        /// Changes reference plane of the nozzle's position plane
        /// Compartment doc must be open.
        /// </summary>
        /// <param name="newRefPlane"></param>
        public void ChangeReferencePlane(Feature newRefPlane)
        {
            try
            {
                //Change refence plane
                FeatureManager.ChangeReferenceOfReferencePlane(newRefPlane, GetPositionPlane());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while changing nozzle's reference plane", ex.Message);
                return;
            }
        }

        /// <summary>
        /// Flips dimension of the nozzle.
        /// Compartment's document must be open
        /// </summary>
        public void FlipDimension()
        {
            Feature positionPlane = GetPositionPlane();
            RefPlaneFeatureData refPlaneFeatureData = positionPlane.GetDefinition();

            // Toggle the dimension
            refPlaneFeatureData.ReversedReferenceDirection[0] = !refPlaneFeatureData.ReversedReferenceDirection[0];

            // Modify the feature within the model
            positionPlane.ModifyDefinition(refPlaneFeatureData, SolidWorksDocumentProvider.GetActiveDoc(), null);
        }

        /// <summary>
        /// Changes distance of the nozzle's position plane from the starting plane.
        /// Compartment document must be open
        /// </summary>
        /// <param name="distance"></param>
        public void ChangeDistance(double distance)
        {
            try
            {
                FeatureManager.ChangeDistanceOfReferencePlane(GetPositionPlane(), distance);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error when changing nozzle position plane's distance.", ex.Message);
            }
        }

        /// <summary>
        /// Deletes a nozzle component from a SolidWorks compartment assembly, including the associated file.
        /// </summary>
        public void DeleteNozzle()
        {
            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();
            Component2 nozzleComponent = GetComponent();

            SelectionMgr selectionManager = (SelectionMgr)compartmentDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the nozzle and its position plane to be deleted
            nozzleComponent.Select4(false, selectData, false);
            GetPositionPlane().Select2(true, 1);

            //Get nozzle document's path to delete the file
            ModelDoc2 componentDocument = nozzleComponent.GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected nozzle
            ((AssemblyDoc)compartmentDoc).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            compartmentDoc.EditRebuild3();

            //Delete the file
            File.Delete(path);
        }

        /// <summary>
        /// Sets a new offset value in meters. Positive value to move nozzle to the right from the middle point, negative - to the left.
        /// </summary>
        /// <param name="distanceInMeters"></param>
        public void SetOffset(double distanceInMeters)
        {
            // Activate Nozzle doc
            ActivateDocument();

            Feature nozzleSketch = GetSketch();

            //Set limit, which is a half of internal diameter.
            double limit = nozzleSketch.Parameter(SHELL_DIAMETER_INTERNAL).Value / 2 / 1000;
            //Set starting point which is a half of external diameter
            double startingPoint = nozzleSketch.Parameter(SHELL_DIAMETER_EXTERNAL).Value / 2 / 1000;

            //Distance must be between limit values to both directions.
            if ((distanceInMeters) <= -limit ||
                (distanceInMeters) >= limit)
            {
                MessageBox.Show($"Incorrect distance value. Distance cannot be < -{limit} and > {limit}");
                return;
            }

            //Set new offset value
            double newOffsetValue = startingPoint - distanceInMeters;
            nozzleSketch.Parameter(OFFSET).SetSystemValue3(
                newOffsetValue, 
                (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting, 
                null);

            ((AssemblyDoc)_currentlyActiveNozzleDoc).EditRebuild();

            CloseDocument();
        }

        /// <summary>
        /// Rotates the nozzle according to its central vertical axis.
        /// </summary>
        /// <param name="angleInDegrees"></param>
        public void RotateNozzle(double angleInDegrees)
        {
            // Activate Nozzle doc
            ActivateDocument();

            // Get nozzle right reference plane
            Feature rightRefPlane = GetNozzleRightRefPlane();

            //Get feature definition
            RefPlaneFeatureData rightRefPlaneFeatData = rightRefPlane.GetDefinition();

            //Set angle before converting it into radians
            rightRefPlaneFeatData.Angle = angleInDegrees * (Math.PI / 180);

            //Modify feature definition
            rightRefPlane.ModifyDefinition(rightRefPlaneFeatData, _currentlyActiveNozzleDoc, null);

            CloseDocument();
        }

        /// <summary>
        /// Rotates nozzle according sketch circle.
        /// </summary>
        /// <param name="angleInDegrees"></param>
        public void SetRotationAngle(double angleInDegrees)
        {
            // Normalize angle to 0-360 degrees (using modulo operator)
            angleInDegrees = (angleInDegrees % 360 + 360) % 360;

            // Activate Nozzle doc
            ActivateDocument();

            try
            {
                //Set angle
                GetSketch().Parameter(CENTER_AXIS_ROTATION_ANGLE)?.SetSystemValue3(
                    angleInDegrees * (Math.PI / 180),
                    (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting,
                    null);

                _currentlyActiveNozzleDoc.EditRebuild3();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting rotation angle: {ex.Message}");
            }

            CloseDocument();
        }

        /// <summary>
        /// Changes distance between nozzle's top plane and sketch external point. This allows the nozzle to be moved up and down.
        /// </summary>
        /// <param name="distanceInMeters"></param>
        public void ChangeNozzleDistance(double distanceInMeters)
        {
            // Activate Nozzle doc
            ActivateDocument();

            // Get top plane mate
            Feature topPlaneMate = GetTopPlaneMate();

            try
            {
                // Change mate distance
                MateManager.ChangeDistance(topPlaneMate, distanceInMeters);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error changing nozzle distance: {ex.Message}");
            }

            // Update attribute with the new PID
            _nozzleSettings.PIDTopPlaneMate = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);

            DocumentManager.UpdateAndSaveDocuments();
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
            DocumentManager.UpdateAndSaveDocuments();
            CloseDocument();
        }

        /// <summary>
        /// Deletes nozzle assembly in the nozzle 
        /// </summary>
        public void DeleteNozzleAssembly()
        {
            ActivateDocument();

            try
            {
                SelectionMgr selectionManager = (SelectionMgr)_currentlyActiveNozzleDoc.SelectionManager;
                SelectData selectData = selectionManager.CreateSelectData();

                //Select the nozzle assembly to be deleted
                GetNozzleAssemblyComp().Select4(false, selectData, false);

                //Delete selected dished end
                ((AssemblyDoc)_currentlyActiveNozzleDoc).DeleteSelections(0);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error while trying to delete nozzle assembly.", ex.Message);
                CloseDocument();
            }
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

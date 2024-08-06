using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
            Feature positionPlane = SWFeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referencePlane,
                distance: distance,
                name: positionPlaneName);

            // Add a new nozzle and make it independent
            Component2 nozzle = ComponentManager.AddComponentAssembly(compartmentDoc, NOZZLE_PATH);
            ComponentManager.MakeComponentIndependent(nozzle, NOZZLE_PATH);

            // Rename nozzle component
            string componentName = $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber-1].Nozzles.Count+1}";
            SWFeatureManager.GetFeatureByName(compartmentDoc, nozzle.Name2).Name = componentName;

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

            void MateNozzle()
            {
                try
                {
                    positionPlaneMate = MateManager.CreateMate(
                        componentFeature1: positionPlane,
                        componentFeature2: SWFeatureManager.GetMajorPlane(nozzle, MajorPlane.Right),
                        alignmentType: MateAlignment.Aligned,
                        name: $"{nozzle.Name2} - {POSITION_PLANE_NAME}");

                    MateManager.CreateMate(
                        componentFeature1: compartmentCenterAxis,
                        componentFeature2: SWFeatureManager.GetFeatureByName(nozzle, "Center Axis"),
                        alignmentType: MateAlignment.Anti_Aligned,
                        name: $"{nozzle.Name2} - {CENTER_AXIS_NAME}");

                    MateManager.CreateMate(
                        componentFeature1: compartmentFrontPlane,
                        componentFeature2: SWFeatureManager.GetMajorPlane(nozzle, MajorPlane.Front),
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
                            "Nozzle axis",
                            "AXIS",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature nozzleAxis = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

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
                            _nozzleSettings.PIDNozzleAxis = nozzleModelDoc.Extension.GetPersistReference3(nozzleAxis);
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

        private void AddCutOutExtrude()
        {
            AddCutOutPlane();

            AddCutOutSketch();

            CreateCutExtrude();
        }

        /// <summary>
        /// Creates a reference plane(a flat construction surface) in the main shell document of a tank assembly.
        /// The reference plane is positioned perpendicular to the axis of the most recently added nozzle and passes 
        /// through the nozzle's midpoint. 
        /// </summary>
        /// <param name="compartmentNumber"></param>
        private void AddCutOutPlane()
        {
            // Get a reference to the main tank assembly in SolidWorks.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // Get references to the specific compartment and the latest nozzle added to it.
            Compartment compartment = tankSiteAssembly._compartmentsManager.Compartments.Last();
            Nozzle nozzle = compartment.Nozzles.Last();

            // Activate the compartment document to make it the active document in SolidWorks.
            tankSiteAssembly._compartmentsManager.ActivateDocument();
            compartment.ActivateDocument();

            // Get the nozzle component object.
            Component2 nozzleComp = nozzle.GetComponent();

            // Activate the nozzle document to access its features.
            nozzle.ActivateDocument();

            // Retrieve the axis and midpoint features of the nozzle.
            Feature nozzleAxis = nozzle.GetNozzleAxis();
            Feature midPoint = nozzle.GetMidPoint();

            // Close the nozzle and compartment documents as we've extracted the needed information.
            nozzle.CloseDocument();
            compartment.CloseDocument();

            // Get a reference to the currently active document, which should be the main shell document.
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Get the SolidWorks component object representing the compartment in the main shell document.
            Component2 compartmentComp = compartment.GetComponent();

            // Ensure nothing is pre-selected to avoid conflicts.
            shellDoc.ClearSelection2(true);

            // Select the nozzle axis and midpoint in the context of the main shell document.
            // This includes references to the compartment and nozzle assembly names for accurate selection.
            shellDoc.Extension.SelectByID2(
                            $"{nozzleAxis.Name}@{compartmentComp.Name2}@{shellDoc.GetTitle()}/{nozzleComp.Name2}@{compartmentComp.Name2.Split('-')[0]}",
                            "AXIS",
                            0, 0, 0,
                            false,
                            0, null, 0);

            shellDoc.Extension.SelectByID2(
                            $"{midPoint.Name}@{compartmentComp.Name2}@{shellDoc.GetTitle()}/{nozzleComp.Name2}@{compartmentComp.Name2.Split('-')[0]}",
                            "DATUMPOINT",
                            0, 0, 0,
                            true,
                            1, null, 0);

            // Get access to the feature manager of the shell document.
            FeatureManager featureManager = shellDoc.FeatureManager;

            // Create a new reference plane that is perpendicular to the nozzle axis and passes through the midpoint.
            Feature cutOutPlane = (Feature)featureManager.InsertRefPlane(
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Perpendicular,
                0,
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident,
                0, 0, 0);

            // Set a descriptive name for the reference plane.
            cutOutPlane.Name = $"{compartmentComp.Name2.Split('-')[0]} {nozzleComp.Name2.Split('-')[0]} Cut out plane";

            // Store the reference plane information in the nozzle settings for later use.
            nozzle._nozzleSettings.PIDCutOutPlane = shellDoc.Extension.GetPersistReference3(cutOutPlane);
        }

        /// <summary>
        /// creates a circular sketch on the main shell document of a tank assembly. 
        /// The circle is centered on the axis of the latest nozzle added to a specific compartment, 
        /// and its radius is determined by the "D1" dimension(representing the diameter) of a "Cutout sketch" 
        /// found within the "Neck" component of the nozzle assembly.This sketch is typically used to define the 
        /// cutout shape for the nozzle on the tank shell.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        private void AddCutOutSketch()
        {
            // 1. Get References to Objects:
            // Retrieve the main tank site assembly object that holds all the tank components.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // Get references to the specific compartment and the latest nozzle added to it.
            Compartment compartment = tankSiteAssembly._compartmentsManager.Compartments.Last();
            Nozzle nozzle = compartment.Nozzles.Last();

            // 2. Activate Documents:
            // Make the compartment document the currently active document in SolidWorks.
            compartment.ActivateDocument();

            // Get the SolidWorks component object that represents the nozzle in the tank site assembly.
            Component2 nozzleComp = nozzle.GetComponent();

            // Activate the nozzle document to gain access to its features and geometry.
            nozzle.ActivateDocument();

            // Get the nozzle axis
            Feature nozzleAxis = GetNozzleAxis();

            // Get the nozzle assembly document, containing all parts of the nozzle assembly.
            ModelDoc2 nozzleAssemblyDoc = nozzle.GetNozzleAssemblyComp().GetModelDoc2();

            // 3. Extract Cutout Radius:
            // Initialize a variable to store the radius of the cutout.
            double radius = 0;

            // Open the nozzle assembly document using a wrapper for convenient access.
            using (var document = new SolidWorksDocumentWrapper(
                SolidWorksDocumentProvider._solidWorksApplication,
               nozzleAssemblyDoc))
            {
                // Iterate through each component within the nozzle assembly.
                foreach (object component in (object[])((AssemblyDoc)nozzleAssemblyDoc).GetComponents(true))
                {
                    // Look for a component named "Neck".
                    if (((Component2)component).Name2.Contains("Neck"))
                    {
                        // Get the "Cutout sketch" feature within the "Neck" component.
                        Feature cutOutScketch = SWFeatureManager.GetFeatureByName((Component2)component, "Cutout sketch");

                        // Get the "D1" dimension from the cutout sketch, which is assumed to represent the diameter.
                        Dimension dimension = cutOutScketch.Parameter("D1");

                        // Extract the dimension value as a double and convert it to a radius.
                        radius = dimension.GetValue3(
                            (int)swInConfigurationOpts_e.swAllConfiguration,
                            null)[0];

                        // Exit the loop since the radius has been found.
                        break;
                    }
                }
            }

            // 4. Create Cutout Sketch on Shell:
            // Close the nozzle and compartment documents, as they are no longer needed.
            nozzle.CloseDocument();
            compartment.CloseDocument();

            // Get the currently active document, which should now be the main shell document.
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Get the SolidWorks component representing the compartment within the shell document.
            Component2 compartmentComp = compartment.GetComponent();

            // Access the sketch manager to manipulate sketches on the shell document.
            SketchManager sketchManager = shellDoc.SketchManager;

            // Select the previously created cutout plane to create the sketch on.
            GetCutOutPlane().Select2(false, 1);

            // Start a new sketch on the selected cutout plane.
            sketchManager.InsertSketch(true);

            // Enter sketch editing mode.
            shellDoc.EditSketch();

            // Create a circle on the sketch with the extracted radius, centered at a default position (0, 0, 0).
            sketchManager.CreateCircleByRadius(0, 0, 0, radius/2000);

            // Select the nozzle axis and the center point of the circle for constraint application.
            shellDoc.Extension.SelectByID2(
                            $"{nozzleAxis.Name}@{compartmentComp.Name2}@{shellDoc.GetTitle()}/{nozzleComp.Name2}@{compartmentComp.Name2.Split('-')[0]}",
                            "AXIS",
                            0, 0, 0,
                            false,
                            0, null, 0);

            shellDoc.Extension.SelectByID2(
                           "Point2",
                           "SKETCHPOINT",
                           0, 0, 0,
                           true,
                           0, null, 0);

            // Add a coincident constraint to align the circle's center with the nozzle axis.
            shellDoc.SketchAddConstraints("sgCOINCIDENT");
        }

        private void CreateCutExtrude()
        {
            // 1. Get References and Setup:
            // Retrieve the main tank site assembly object.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // Get references to the specific compartment and the latest nozzle added to it.
            Compartment compartment = tankSiteAssembly._compartmentsManager.Compartments.Last();
            Nozzle nozzle = compartment.Nozzles.Last();

            // Get the main shell document where the cut-extrude will be applied.
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();
            FeatureManager featureManager = shellDoc.FeatureManager;

            // Select all cylindrical shells in the assembly for the cut-extrude operation.
            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.SelectAllCylindricalShells();

            // 2.Temporarily Suppress Nozzle Assembly:
            // Activate the compartment document to work with it
             compartment.ActivateDocument();

            // Activate the nozzle document to access its components.
             nozzle.ActivateDocument();

           // Get the SolidWorks component object representing the nozzle assembly.
           Component2 nozzleAssemblyComp = nozzle.GetNozzleAssemblyComp();

            // Close the nozzle and compartment documents after obtaining the nozzle assembly component.
             nozzle.CloseDocument();
            compartment.CloseDocument();

           // Suppress the nozzle assembly to avoid it being affected by the cut-extrude operation.
           SWFeatureManager.Suppress(nozzleAssemblyComp);

            // 3. Create Cut-Extrude Feature:
            // Perform the cut-extrude operation on the selected cylindrical shells.
            Feature cutExtrude = (Feature)featureManager.FeatureCut4(
                true,
                false,
                false,
                (int)swEndConditions_e.swEndCondThroughAll,
                (int)swEndConditions_e.swEndCondBlind,
                0.01,
                0.01,
                false,
                false,
                false,
                false,
                0,
                0,
                false,
                false,
                false,
                false,
                false,
                false,
                true,
                false,
                true,
                true,
                (int)swStartConditions_e.swStartSketchPlane,
                0,
                false,
                false);

            // 4. Unsuppress Nozzle Assembly and Save:
            // Unsuppress the nozzle assembly, making it visible again.
            SWFeatureManager.Unsuppress(nozzleAssemblyComp);

            // Store a reference to the newly created cut-extrude feature in the nozzle settings for later use.
            nozzle._nozzleSettings.PIDCutExtrude = shellDoc.Extension.GetPersistReference3(cutExtrude);

            // Update the SolidWorks documents to reflect the changes and save them.
            DocumentManager.UpdateAndSaveDocuments();
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
                SWFeatureManager.ChangeReferenceOfReferencePlane(newRefPlane, GetPositionPlane());
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
                SWFeatureManager.ChangeDistanceOfReferencePlane(GetPositionPlane(), distance);
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
        public void AddNozzleAssembly(int compartmentNumber)
    {
        // Activate current nozzle document
            ActivateDocument();

            // Add nozzle assembly and make it independent
            Component2 nozzleAssembly = ComponentManager.AddComponentAssembly(_currentlyActiveNozzleDoc, NOZZLE_ASSEMBLY_PATH);
            ComponentManager.MakeComponentIndependent(nozzleAssembly, NOZZLE_ASSEMBLY_PATH);

            // Get features for mating
            Feature nozzleAssemblyCenterAxis = SWFeatureManager.GetFeatureByName(nozzleAssembly, "Center axis");

            // Create Mates
            try
            {
                MateManager.CreateMate(
                    componentFeature1: SWFeatureManager.GetFeatureByName(_currentlyActiveNozzleDoc, "Nozzle axis"),
                    componentFeature2: nozzleAssemblyCenterAxis,
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {CENTER_AXIS_NAME}");

                MateManager.CreateMate(
                    componentFeature1: GetNozzleRightRefPlane(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {RIGHT_PLANE_NAME}");

                Feature topPlaneMate = MateManager.CreateMate(
                    componentFeature1: GetCutPlane(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Top),
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

            AddCutOutExtrude();

            DocumentManager.UpdateAndSaveDocuments();
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

        public Feature GetNozzleAxis() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                      _nozzleSettings.PIDNozzleAxis,
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

        public Feature GetCutOutPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDCutOutPlane,
                     out int error);

        public Feature GetCutExtrude() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDCutExtrude,
                     out int error);
    }
}

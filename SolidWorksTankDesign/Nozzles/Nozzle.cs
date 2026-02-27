using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;

namespace SolidWorksTankDesign
{
    public class Nozzle
    {
        private ModelDoc2 _currentlyActiveNozzleDoc;

        private const string MANHOLE_NAME = "M";
        private const string POSITION_PLANE_NAME = "position plane";
        private const string CENTER_AXIS_NAME = "Center axis";
        private const string FRONT_PLANE_NAME = "Front plane";
        private const string RIGHT_PLANE_NAME = "Right plane";
        private const string TOP_PLANE_NAME = "Top plane";
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
        /// Changes external and internal diameters, and nozzle's offset of nozzle position sketch
        /// </summary>
        /// <param name="nozzleDoc"></param>
        /// <param name="externalDiameter"></param>
        private void ChangeNozzleSketchDiametersAndNozzleOffset(ModelDoc2 nozzleDoc, double externalDiameter)
        {
            // Get the sketch
            Feature scketch = SWFeatureManager.GetFeatureByName(nozzleDoc, "Sketch");

            // Get external and internal diameters, and nozzle's offset dimensions
            Dimension externalDiameterDimension = scketch.Parameter("ShellDiameterExternal");
            Dimension internalDiameterDimension = scketch.Parameter("ShellDiameterInternal");
            Dimension nozzleOffset = scketch.Parameter("Offset");

            // Set new values
            externalDiameterDimension.SetValue3(externalDiameter, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
            internalDiameterDimension.SetValue3(externalDiameter - 50, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
            nozzleOffset.SetValue3((externalDiameter - 50)/2, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
        }

        /// <summary>
        /// This constructor is responsible for creating and positioning a manhole within a SolidWorks assembly.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        /// <param name="referencePlane"></param>
        /// <param name="compartmentCenterAxis"></param>
        /// <param name="compartmentFrontPlane"></param>
        /// <param name="distance"></param>
        public Nozzle(
            string compartmentFolder,
            string nozzlePositionSketchPath,
            int compartmentNumber,
            Feature referencePlane,
            Feature compartmentCenterAxis,
            Feature compartmentFrontPlane,
            double distance,
            bool flip,
            double externalDiameter)
        {
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;

            // Create a path where the empty manhole doc will be saved
            string ticks = DateTime.Now.Ticks.ToString();
            string targetPath = Path.Combine(compartmentFolder, 
                $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber].Nozzles.Count + 1}_{ticks}.SLDASM");

            // Open empty manhole doc and save it to a new destination
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(nozzlePositionSketchPath);
            ModelDoc2 nozzlePositionSketchDoc = solidWorksApp.OpenDoc7(documentSpecification);

            nozzlePositionSketchDoc.SaveAs3(targetPath, 0, 0);

            // Close empty manhole and newly saved docs
            solidWorksApp.CloseDoc(nozzlePositionSketchDoc.GetTitle());
            solidWorksApp.CloseDoc(targetPath);

            // Get active compartment's document
            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();

            string positionPlaneName = $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber].Nozzles.Count + 1} {POSITION_PLANE_NAME}";
            // Create manhole's position plane
            Feature positionPlane = SWFeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referencePlane,
                distance: distance,
                name: positionPlaneName,
                flip: flip);

            // Add a new manhole
            Component2 nozzle = ComponentManager.AddComponentAssembly(compartmentDoc, targetPath);

            ModelDoc2 nozzleModelDoc = nozzle.GetModelDoc2();

            ChangeNozzleSketchDiametersAndNozzleOffset(nozzleModelDoc, externalDiameter);

            Feature positionPlaneMate = null;
            MateNozzle();

            // Get manhole Entities and Initialize Settings
            _nozzleSettings = new NozzleSettings();
            try
            {
                GetNozzlePIDs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting manhole entities: {ex.Message}");
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
                    MessageBox.Show(ex.Message, "At least one of manhole mates could not be created.");
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

                    // Use a SolidWorksDocumentWrapper for managing the manhole's model document.
                    using (var nozzleDocument = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, nozzleModelDoc))
                    {
                        // Get the selection manager to interact with selections within the manhole's model
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
                            // Populate the _nozzleSettings with the retrieved PIDs for those entities that has to be reachable from manhole's document
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
                    MessageBox.Show(ex.Message, "Could not initialize manhole settings.");
                    return;
                }
            }
        }

        /// <summary>
        /// Attempts to create a cutout(hole or removal of material) in the active SolidWorks document
        /// </summary>
        public void AddCutOutExtrude(Compartment compartment, Nozzle nozzle)
        {
            try
            {
                // Create cut out plane
                AddCutOutPlane(compartment, nozzle);

                // Get sketch name
                string sketchName = AddCutOutSketch(compartment, nozzle);

                // Create cut
                CreateCutExtrude(sketchName, compartment, nozzle);
            }
           catch(Exception ex)
            {
                MessageBox.Show("Error while adding cut extrude.", ex.Message);
            }
        }

        /// <summary>
        /// Creates a reference plane(a flat construction surface) in the main shell document of a tank assembly.
        /// The reference plane is positioned perpendicular to the axis of the most recently added manhole and passes 
        /// through the manhole's midpoint. 
        /// </summary>
        /// <param name="compartmentNumber"></param>
        private void AddCutOutPlane(Compartment compartment, Nozzle nozzle)
        {
            // Get a reference to the main tank assembly in SolidWorks.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // Activate the compartment document to make it the active document in SolidWorks.
            tankSiteAssembly._compartmentsManager.ActivateDocument();
            compartment.ActivateDocument();

            // Get the manhole component object.
            Component2 nozzleComp = nozzle.GetComponent();

            // Activate the manhole document to access its features.
            nozzle.ActivateDocument();

            // Retrieve the axis and midpoint features of the manhole.
            Feature nozzleAxis = nozzle.GetNozzleAxis();
            Feature midPoint = nozzle.GetMidPoint();

            // Close the manhole and compartment documents as we've extracted the needed information.
            nozzle.CloseDocument();
            compartment.CloseDocument();

            // Get a reference to the currently active document, which should be the main shell document.
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Get the SolidWorks component object representing the compartment in the main shell document.
            Component2 compartmentComp = compartment.GetComponent();

            // Ensure nothing is pre-selected to avoid conflicts.
            shellDoc.ClearSelection2(true);

            // Select the manhole axis and midpoint in the context of the main shell document.
            // This includes references to the compartment and manhole assembly names for accurate selection.
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

            // Create a new reference plane that is perpendicular to the manhole axis and passes through the midpoint.
            Feature cutOutPlane = (Feature)featureManager.InsertRefPlane(
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Perpendicular,
                0,
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident,
                0, 0, 0);

            // Set a descriptive name for the reference plane.
            cutOutPlane.Name = $"{compartmentComp.Name2.Split('-')[0]} {nozzleComp.Name2.Split('-')[0]} Cut out plane";

            // Store the reference plane information in the manhole settings for later use.
            nozzle._nozzleSettings.PIDCutOutPlane = shellDoc.Extension.GetPersistReference3(cutOutPlane);
        }

        /// <summary>
        /// creates a circular sketch on the main shell document of a tank assembly. 
        /// The circle is centered on the axis of the latest manhole added to a specific compartment, 
        /// and its radius is determined by the "D1" externalDiameterDimension(representing the diameter) of a "Cutout sketch" 
        /// found within the "Neck" component of the manhole assembly.This sketch is typically used to define the 
        /// cutout shape for the manhole on the tank shell.
        /// </summary>
        /// <param name="compartmentNumber"></param>
        private string AddCutOutSketch(Compartment compartment, Nozzle nozzle)
        {
            // 1. Get References to Objects:
            // Retrieve the main tank site assembly object that holds all the tank components.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // 2. Activate Documents:
            // Make the compartment document the currently active document in SolidWorks.
            compartment.ActivateDocument();

            // Get the SolidWorks component object that represents the manhole in the tank site assembly.
            Component2 nozzleComp = nozzle.GetComponent();

            // Activate the manhole document to gain access to its features and geometry.
            nozzle.ActivateDocument();

            // Get the manhole axis
            Feature nozzleAxis = GetNozzleAxis();

            // Get the manhole assembly component, containing all parts of the manhole assembly.
            Component2 nozzleAssemblyComp = nozzle.GetNozzleAssemblyComp();

            // 3. Extract Cutout Radius:

            // Get the "Cutout sketch" feature within the manhole assembly component.
            Feature cutOutScketch = SWFeatureManager.GetFeatureByName(nozzleAssemblyComp, "Cut out sketch");

            // Get the "D1" externalDiameterDimension from the cutout sketch, which is assumed to represent the diameter.
            Dimension dimension = cutOutScketch.Parameter("D1");

            // Extract the externalDiameterDimension value as a double and convert it to a radius.
            double radius = dimension.GetValue3(
                (int)swInConfigurationOpts_e.swAllConfiguration,
                null)[0];

            // 4. Create Cutout Sketch on Shell:
            // Close the manhole and compartment documents, as they are no longer needed.
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

            // Select the manhole axis and the center point of the circle for constraint application.
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

            // Add a coincident constraint to align the circle's center with the manhole axis.
            shellDoc.SketchAddConstraints("sgCOINCIDENT");

            // Get active sketch
            Sketch sketch = shellDoc.GetActiveSketch2();

            return ((Feature)sketch).Name;
        }

        /// <summary>
        /// Creates a cut-out in the cylindrical shells of a tank assembly in SolidWorks. 
        /// The cut-out is specifically designed to accommodate a manhole that has been recently added to the tank.
        /// </summary>
        private void CreateCutExtrude(string sketchName, Compartment compartment, Nozzle nozzle)
        {
            // 1. Get References and Setup:
            // Retrieve the main tank site assembly object.
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            // Get the main shell document where the cut-extrude will be applied.
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();
            FeatureManager featureManager = shellDoc.FeatureManager;

            //2.Create Cut - Extrude Feature:
            //Perform the cut - extrude operation on the selected cylindrical shells.
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
                true,
                false,
                true,
                false,
                false,
                (int)swStartConditions_e.swStartSketchPlane,
                0,
                false,
                false);

            if (cutExtrude == null)
            {
                MessageBox.Show("Cut extrude was not created.");
            }

            ExtrudeFeatureData2 cutExtrudeFeatData = cutExtrude.GetDefinition();

            cutExtrude.Select2(false, 1);

            foreach (CylindricalShell cylindricalShell in tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells)
            {
                ModelDoc2 assemblyOfCylindricalShellsDoc = tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();
                string cylindricalShellName = cylindricalShell.GetComponent().Name2;

                tankSiteAssembly._assemblyOfCylindricalShells.CloseDocument();

                ((AssemblyDoc)shellDoc).AddToFeatureScope($"{assemblyOfCylindricalShellsDoc.GetTitle()}-1@{shellDoc.GetTitle()}/{cylindricalShellName}@{assemblyOfCylindricalShellsDoc.GetTitle()}");
                ((AssemblyDoc)shellDoc).UpdateFeatureScope();
            }


            shellDoc.ClearSelection2(true);

            cutExtrude.ModifyDefinition(cutExtrudeFeatData, shellDoc, null);

            // Store a reference to the newly created cut-extrude feature in the manhole settings for later use.
            nozzle._nozzleSettings.PIDCutExtrude = shellDoc.Extension.GetPersistReference3(cutExtrude);

            // Update the SolidWorks documents to reflect the changes and save them.
            DocumentManager.UpdateAndSaveDocuments();
        }

        /// <summary>
        /// Changes reference plane of the manhole's position plane
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
                MessageBox.Show("Error while changing manhole's reference plane", ex.Message);
                return;
            }
        }

        /// <summary>
        /// Flips externalDiameterDimension of the manhole.
        /// Compartment's document must be open
        /// </summary>
        public void FlipDimension()
        {
            Feature positionPlane = GetPositionPlane();
            RefPlaneFeatureData refPlaneFeatureData = positionPlane.GetDefinition();

            // Toggle the externalDiameterDimension
            refPlaneFeatureData.ReversedReferenceDirection[0] = !refPlaneFeatureData.ReversedReferenceDirection[0];

            // Modify the feature within the model
            positionPlane.ModifyDefinition(refPlaneFeatureData, SolidWorksDocumentProvider.GetActiveDoc(), null);
        }

        /// <summary>
        /// Changes distance of the manhole's position plane from the starting plane.
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
                MessageBox.Show("Error when changing manhole position plane's distance.", ex.Message);
            }
        }

        /// <summary>
        /// Deletes cut by deleting its cut out plane
        /// </summary>
        /// <returns></returns>
        public bool DeleteCutExtrude()
        {
            ModelDoc2 shellDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Get cut out plane
            Feature cutOutPlane = GetCutOutPlane();

            if (cutOutPlane == null) return true;

            // Select cut out plane
            GetCutOutPlane().Select2(false, 1);

            // Delete cut out plane
            shellDoc.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Children);

            return true;
        }

        /// <summary>
        /// Deletes a manhole component from a SolidWorks compartment assembly, including the associated file.
        /// </summary>
        public void DeleteNozzle()
        {
            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();
            Component2 nozzleComponent = GetComponent();

            SelectionMgr selectionManager = (SelectionMgr)compartmentDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the manhole and its position plane to be deleted
            nozzleComponent.Select4(false, selectData, false);
            GetPositionPlane().Select2(true, 1);

            //Get manhole document's path to delete the file
            ModelDoc2 componentDocument = nozzleComponent.GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected manhole
            ((AssemblyDoc)compartmentDoc).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            compartmentDoc.EditRebuild3();

            //Delete the file
            System.IO.File.Delete(path);
        }

        /// <summary>
        /// Sets a new offset value in meters. Positive value to move manhole to the right from the middle point, negative - to the left.
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
        /// Rotates the manhole according to its central vertical axis.
        /// </summary>
        /// <param name="angleInDegrees"></param>
        public void RotateNozzle(double angleInDegrees)
        {
            // Activate Nozzle doc
            ActivateDocument();

            // Get manhole right reference plane
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
        /// Rotates manhole according sketch circle.
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
        /// Changes distance between manhole's top plane and sketch external point. This allows the manhole to be moved up and down.
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
                MessageBox.Show($"Error changing manhole distance: {ex.Message}");
            }

            // Update attribute with the new PID
            _nozzleSettings.PIDTopPlaneMate = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);

            DocumentManager.UpdateAndSaveDocuments();
        }

        /// <summary>
        /// Inserts a pre-designed manhole assembly into the currently active SolidWorks document. 
        /// It first prepares the manhole assembly by packaging it using the PackAndGo functionality. 
        /// Then, it adds the packaged assembly to the active document and precisely positions it using mates (geometric constraints)
        /// that align key features of the assembly with corresponding features in the active document.
        /// Finally, it creates a cutout to accommodate the newly added manhole assembly and saves the modified document.
        /// </summary>
        public void AddNozzleAssembly(string compartmentFolder, string nozzleDocPath, Compartment compartment, Nozzle nozzle)
        {
            // Activate current manhole document
            ActivateDocument();

            int error = 0;
            int warning = 0;

            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
            DocumentSpecification documentSpecification = (DocumentSpecification)solidWorksApp.GetOpenDocSpec(nozzleDocPath);
            documentSpecification.Silent = true;
            ModelDoc2 nozzleAssemblyDoc = solidWorksApp.OpenDoc7(documentSpecification);

            if(nozzleAssemblyDoc is null)
            {
                string nozzleDocTitle = nozzleDocPath.Split('\\').Last().Split('.')[0];
                object[] activeDocs = solidWorksApp.GetDocuments();
                foreach (object activeDoc in activeDocs)
                {
                    string nameDoc = ((ModelDoc2)activeDoc).GetTitle();
                    if (nameDoc == nozzleDocTitle)
                    {
                        nozzleAssemblyDoc = (ModelDoc2)activeDoc;
                        break;
                    }
                }
            }

            //string nameOfCurrentDoc = SolidWorksDocumentProvider.GetActiveDoc().GetTitle();
            string name = nozzleAssemblyDoc.GetTitle();

            // Package the manhole assembly and its associated files using Pack and Go, and get the path to the packed assembly
            string path = DocumentManager.PackAndGo(compartmentFolder, nozzleAssemblyDoc, null, null);

            // Close the manhole assembly document after it has been packed
            solidWorksApp.CloseDoc(nozzleAssemblyDoc.GetTitle());
            string docpath = _currentlyActiveNozzleDoc.GetPathName();

            // Add the packed manhole assembly to the currently active manhole document as a component
            Component2 nozzleAssembly = ComponentManager.AddComponentAssembly(_currentlyActiveNozzleDoc, path);

            // Get a reference to the "Center axis" feature of the added manhole assembly, which will be used for mating
            Feature nozzleAssemblyCenterAxis = SWFeatureManager.GetFeatureByName(nozzleAssembly, "Center axis");

            try
            {
                // Create mates to position and align the manhole assembly within the active document
                // 1. Align the "Nozzle axis" of the active document with the "Center axis" of the manhole assembly
                MateManager.CreateMate(
                    componentFeature1: SWFeatureManager.GetFeatureByName(_currentlyActiveNozzleDoc, "Nozzle axis"),
                    componentFeature2: nozzleAssemblyCenterAxis,
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {CENTER_AXIS_NAME}");

                // 2. Align the right plane of the active manhole with the right plane of the manhole assembly
                MateManager.CreateMate(
                    componentFeature1: GetNozzleRightRefPlane(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Right),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {RIGHT_PLANE_NAME}");

                // 3. Anti-align the top plane of the manhole assembly with a "Cut plane" in the active document
                Feature topPlaneMate = MateManager.CreateMate(
                    componentFeature1: GetCutPlane(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Top),
                    alignmentType: MateAlignment.Anti_Aligned,
                    distance: 0,
                    name: $"{nozzleAssembly.Name2} - {TOP_PLANE_NAME}");

                // Store persistent references (PIDs) to the manhole assembly component and the top plane mate for future use
                _nozzleSettings.PIDNozzleAssemblyComp = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(nozzleAssembly);
                _nozzleSettings.PIDTopPlaneMate = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "At least one of manhole assembly mates could not be created.");
            }

            // Update and save the active document and any associated attribute documents
            DocumentManager.UpdateAndSaveDocuments();

            // Add a cutout extrude feature (presumably to create space for the manhole assembly)
            AddCutOutExtrude(compartment, nozzle);

            // Update and save the active document and any associated attribute documents
            DocumentManager.UpdateAndSaveDocuments();
        }

        /// <summary>
        /// Deletes manhole assembly in the manhole 
        /// </summary>
        public void DeleteNozzleAssembly()
        {
            ActivateDocument();

            try
            {
                SelectionMgr selectionManager = (SelectionMgr)_currentlyActiveNozzleDoc.SelectionManager;
                SelectData selectData = selectionManager.CreateSelectData();

                //Select the manhole assembly to be deleted
                GetNozzleAssemblyComp().Select4(false, selectData, false);

                //Delete selected dished end
                ((AssemblyDoc)_currentlyActiveNozzleDoc).DeleteSelections(0);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error while trying to delete manhole assembly.", ex.Message);
                CloseDocument();
            }

            CloseDocument();
        }

        /// <summary>
        /// Activates document of manhole assembly
        /// </summary>
        public void ActivateDocument()
        {
            // Get compartment doc
            ModelDoc2 nozzleModelDoc = GetComponent().GetModelDoc2();

            // Activate compartment doc
            _currentlyActiveNozzleDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(
                nozzleModelDoc.GetTitle() + ".sldasm", 
                true, 
                0, 
                0);
        }

        /// <summary>
        /// Closes active document of manhole assembly
        /// </summary>
        public void CloseDocument()
        {
            if (_currentlyActiveNozzleDoc == null) return;

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(_currentlyActiveNozzleDoc.GetTitle());
            _currentlyActiveNozzleDoc = null;
        }

        public Nozzle DeepClone()
        {
            // Create a new instance of Nozzle
            var clonedNozzle = new Nozzle();

            // Clone the NozzleSettings if it's not null
            if (this._nozzleSettings != null)
            {
                clonedNozzle._nozzleSettings = this._nozzleSettings.DeepClone();
            }

            return clonedNozzle;
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

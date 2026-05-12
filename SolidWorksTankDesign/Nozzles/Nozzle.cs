using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Xml.Linq;
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
        private const double cylinderWallThickness = 50;
        private const double CUT_CLEARANCE_MM = 2.0;

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
            internalDiameterDimension.SetValue3(externalDiameter - cylinderWallThickness, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
            // Neutral offset = externalRadius, matching the startingPoint reference in SetOffset (user offset 0 → externalRadius).
            nozzleOffset.SetValue3(externalDiameter / 2, (int)swSetValueInConfiguration_e.swSetValue_InAllConfigurations, "");
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

            // Open the template nozzle position sketch doc
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(nozzlePositionSketchPath);
            documentSpecification.Silent = true;
            ModelDoc2 nozzlePositionSketchDoc = solidWorksApp.OpenDoc7(documentSpecification);

            if (nozzlePositionSketchDoc is null)
            {
                string templateTitle = System.IO.Path.GetFileNameWithoutExtension(nozzlePositionSketchPath);
                object[] activeDocs = solidWorksApp.GetDocuments();

                foreach (object activeDoc in activeDocs)
                {
                    string nameDoc = ((ModelDoc2)activeDoc).GetTitle();
                    if (nameDoc == templateTitle)
                    {
                        nozzlePositionSketchDoc = (ModelDoc2)activeDoc;
                        break;
                    }
                }
            }

            string name = nozzlePositionSketchDoc.GetTitle();

            int nozzleNumber = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager
                .Compartments[compartmentNumber].Nozzles.Count + 1;

            // Pack and Go — all files get a consistent ticks prefix, preserving all internal
            // in-context references. The generated filename cannot be changed after packing
            // without breaking the envelope part's in-context relations.
            string path = DocumentManager.PackAndGoManhole(
                SolidWorksDocumentProvider.ProjectFolderPath,
                nozzlePositionSketchDoc,
                nozzleNumber);

            // Close the template doc — the packed copy at 'path' is used from here on
            solidWorksApp.CloseDoc(nozzlePositionSketchDoc.GetTitle());

            // Get active compartment's document
            ModelDoc2 compartmentDoc = SolidWorksDocumentProvider.GetActiveDoc();

            string positionPlaneName = $"{MANHOLE_NAME}{SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber].Nozzles.Count + 1} {POSITION_PLANE_NAME}";
            // Create manhole's position plane
            Feature positionPlane = SWFeatureManager.CreateReferencePlaneWithDistance(
                existingPlane: referencePlane,
                distance: distance,
                name: positionPlaneName,
                flip: flip);

            // Add the nozzle using the Pack and Go output path
            Component2 nozzle = ComponentManager.AddComponentAssembly(compartmentDoc, path);

            ModelDoc2 nozzleModelDoc = nozzle.GetModelDoc2();

            ChangeNozzleSketchDiametersAndNozzleOffset(nozzleModelDoc, externalDiameter);

            Feature positionPlaneMate = null;
            Feature rotationMate = null;
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

                    rotationMate = MateManager.CreateMate(
                        externalEntity: (Entity)compartmentFrontPlane,
                        componentEntity: (Entity)SWFeatureManager.GetMajorPlane(nozzle, MajorPlane.Front),
                        referenceEntity: (Entity)compartmentCenterAxis,
                        angle: 0,
                        flipDimension: true,
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
                    _nozzleSettings.PIDRotationMate = compartmentDoc.Extension.GetPersistReference3(rotationMate);

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
                            "Top point",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature topPoint = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

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

                        nozzleModelDoc.Extension.SelectByID2(
                            "Centerline wall intersection",
                            "DATUMPOINT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature centerlineWallIntersection = selectionMgrAtNozzle.GetSelectedObject6(1, -1);

                        try
                        {
                            // Populate the _nozzleSettings with the retrieved PIDs for those entities that has to be reachable from manhole's document
                            _nozzleSettings.PIDCenterAxis = nozzleModelDoc.Extension.GetPersistReference3(nozzleCenterAxis);
                            _nozzleSettings.PIDNozzleAxis = nozzleModelDoc.Extension.GetPersistReference3(nozzleAxis);
                            _nozzleSettings.PIDExternalPoint = nozzleModelDoc.Extension.GetPersistReference3(externalPoint);
                            _nozzleSettings.PIDInternalPoint = nozzleModelDoc.Extension.GetPersistReference3(internalPoint);
                            _nozzleSettings.PIDInsidePoint = nozzleModelDoc.Extension.GetPersistReference3(insidePoint);
                            _nozzleSettings.PIDMidPoint = nozzleModelDoc.Extension.GetPersistReference3(midPoint);
                            _nozzleSettings.PIDTopPoint = nozzleModelDoc.Extension.GetPersistReference3(topPoint);
                            _nozzleSettings.PIDNozzleRightRefPlane = nozzleModelDoc.Extension.GetPersistReference3(nozzleRightRefPlane);
                            _nozzleSettings.PIDCutPlane = nozzleModelDoc.Extension.GetPersistReference3(plane1);
                            _nozzleSettings.PIDSketch = nozzleModelDoc.Extension.GetPersistReference3(sketch);
                            _nozzleSettings.PIDCenterlineWallIntersection = nozzleModelDoc.Extension.GetPersistReference3(centerlineWallIntersection);
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
        /// Inserts a pre-designed manhole assembly into the currently active SolidWorks document. 
        /// It first prepares the manhole assembly by packaging it using the PackAndGo functionality. 
        /// Then, it adds the packaged assembly to the active document and precisely positions it using mates (geometric constraints)
        /// that align key features of the assembly with corresponding features in the active document.
        /// Finally, it creates a cutout to accommodate the newly added manhole assembly and saves the modified document.
        /// </summary>
        public void AddNozzleAssembly(string nozzleDocPath, Compartment compartment, Nozzle nozzle)
        {
            // Activate current manhole document
            ActivateDocument();

            int error = 0;
            int warning = 0;

            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
            DocumentSpecification documentSpecification = (DocumentSpecification)solidWorksApp.GetOpenDocSpec(nozzleDocPath);
            documentSpecification.Silent = true;
            ModelDoc2 nozzleAssemblyDoc = solidWorksApp.OpenDoc7(documentSpecification);

            if (nozzleAssemblyDoc is null)
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
            string path = DocumentManager.PackAndGo(SolidWorksDocumentProvider.ProjectFolderPath, nozzleAssemblyDoc, null, null);

            // Close the manhole assembly document after it has been packed
            solidWorksApp.CloseDoc(nozzleAssemblyDoc.GetTitle());
            string docpath = _currentlyActiveNozzleDoc.GetPathName();

            // Add the packed manhole assembly to the currently active manhole document as a component
            Component2 nozzleAssembly = ComponentManager.AddComponentAssembly(_currentlyActiveNozzleDoc, path);

            // Build _nozzleSettings.NozzleAssemblyComponents list from the sub-components of the nozzle assembly
            _nozzleSettings.NozzleAssemblyComponents.Clear();
            object[] subComponents = (object[])((AssemblyDoc)nozzleAssembly.GetModelDoc2()).GetComponents(true);
            if (subComponents != null)
            {
                SelectionMgr selMgr = (SelectionMgr)_currentlyActiveNozzleDoc.SelectionManager;
                string nozzleDocTitle = _currentlyActiveNozzleDoc.GetTitle();
                string nozzleAssemblyDocTitle = System.IO.Path.GetFileNameWithoutExtension(path);

                foreach (object obj in subComponents)
                {
                    Component2 subComp = (Component2)obj;
                    NozzleAssemblyComponent assemblyComponent = new NozzleAssemblyComponent();

                    // Read ComponentType and HasAdjustableLength from SolidWorks custom properties
                    string componentTypeStr = SWFeatureManager.GetCustomPropertyFromComponent(subComp, "ComponentType");
                    if (Enum.TryParse(componentTypeStr, out NozzleComponentType componentType))
                        assemblyComponent.Settings.ComponentType = componentType;

                    string hasAdjustableLengthStr = SWFeatureManager.GetCustomPropertyFromComponent(subComp, "HasAdjustableLength");
                    if (bool.TryParse(hasAdjustableLengthStr, out bool hasAdjustableLength))
                        assemblyComponent.Settings.HasAdjustableLength = hasAdjustableLength;

                    // Store component PID by selecting it in _currentlyActiveNozzleDoc context
                    // Path format: "SubAssemblyInstance@NozzleDocTitle/SubCompInstance@SubAssemblyDocTitle"
                    string componentSelPath = $"{nozzleAssembly.Name2}@{nozzleDocTitle}/{subComp.Name2}@{nozzleAssemblyDocTitle}";
                    bool componentSelected = _currentlyActiveNozzleDoc.Extension.SelectByID2(componentSelPath, "COMPONENT", 0, 0, 0, false, 0, null, 0);
                    if (componentSelected)
                    {
                        Component2 selectedComp = (Component2)selMgr.GetSelectedObject6(1, -1);
                        if (selectedComp != null)
                            assemblyComponent.Settings.PIDComponent = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(selectedComp);
                    }

                    // For Flange: Plane1 = mating, Plane2 = free
                    // For all other types: Plane2 = mating, Plane1 = free
                    string matingPlaneName = assemblyComponent.Settings.ComponentType == NozzleComponentType.Flange ? "Plane1" : "Plane2";
                    string freePlaneName = assemblyComponent.Settings.ComponentType == NozzleComponentType.Flange ? "Plane2" : "Plane1";

                    // Store mating plane PID by selecting it in _currentlyActiveNozzleDoc context
                    // Path format: "PlaneName@SubAssemblyInstance@NozzleDocTitle/SubCompInstance@SubAssemblyDocTitle"
                    string matingPlanePath = $"{matingPlaneName}@{nozzleAssembly.Name2}@{nozzleDocTitle}/{subComp.Name2}@{nozzleAssemblyDocTitle}";
                    bool matingSelected = _currentlyActiveNozzleDoc.Extension.SelectByID2(matingPlanePath, "PLANE", 0, 0, 0, false, 0, null, 0);
                    if (matingSelected)
                    {
                        Feature matingPlane = (Feature)selMgr.GetSelectedObject6(1, -1);
                        if (matingPlane != null)
                            assemblyComponent.Settings.PIDMatingPlane = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(matingPlane);
                    }

                    // Store free plane PID by selecting it in _currentlyActiveNozzleDoc context
                    string freePlanePath = $"{freePlaneName}@{nozzleAssembly.Name2}@{nozzleDocTitle}/{subComp.Name2}@{nozzleAssemblyDocTitle}";
                    bool freeSelected = _currentlyActiveNozzleDoc.Extension.SelectByID2(freePlanePath, "PLANE", 0, 0, 0, false, 0, null, 0);
                    if (freeSelected)
                    {
                        Feature freePlane = (Feature)selMgr.GetSelectedObject6(1, -1);
                        if (freePlane != null)
                            assemblyComponent.Settings.PIDFreePlane = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(freePlane);
                    }

                    _nozzleSettings.NozzleAssemblyComponents.Add(assemblyComponent);
                }

                // Sort components top to bottom by distance from Inside point
                _nozzleSettings.NozzleAssemblyComponents = GetComponentsSortedTopToBottom();
            }

            // Get a reference to the "Center axis" feature of the added manhole assembly, which will be used for mating
            Feature nozzleAssemblyCenterAxis = SWFeatureManager.GetFeatureByName(nozzleAssembly, "Center axis");

            try
            {
                // Create mates to position and align the manhole assembly within the active document
                // 1. Align the "Nozzle axis" of the active document with the "Center axis" of the manhole assembly
                Feature axisMate = MateManager.CreateMate(
                    componentFeature1: SWFeatureManager.GetFeatureByName(_currentlyActiveNozzleDoc, "Nozzle axis"),
                    componentFeature2: nozzleAssemblyCenterAxis,
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {CENTER_AXIS_NAME}");

                // 2. Align the right plane of the active manhole with the right plane of the manhole assembly
                MateManager.CreateMate(
                    componentFeature1: GetNozzleRightRefPlane(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Top),
                    alignmentType: MateAlignment.Aligned,
                    name: $"{nozzleAssembly.Name2} - {RIGHT_PLANE_NAME}");

                // 3. Anti-align the top plane of the manhole assembly with a "Cut plane" in the active document
                Feature topPlaneMate = MateManager.CreateMate(
                    componentFeature1: GetTopPoint(),
                    componentFeature2: SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Right),
                    alignmentType: MateAlignment.Anti_Aligned,
                    distance: 0,
                    name: $"{nozzleAssembly.Name2} - {TOP_PLANE_NAME}");

                // Store persistent references (PIDs) to the manhole assembly component and the top plane mate for future use
                _nozzleSettings.PIDNozzleAssemblyComp = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(nozzleAssembly);
                _nozzleSettings.PIDTopPlaneMate = _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);

                // After mates are created, check if the Right plane is pointing downward and flip the axis mate if needed
                Feature nozzleAssemblyRightPlane = SWFeatureManager.GetMajorPlane(nozzleAssembly, MajorPlane.Right);
                bool isPlanePointingUp = SWFeatureManager.IsPlaneNormalPointingUp(nozzleAssemblyRightPlane);

                // If the plane is pointing down (not up), flip the axis mate
                if (!isPlanePointingUp && axisMate != null)
                {
                    MateManager.FlipMate(axisMate);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "At least one of manhole assembly mates could not be created.");
            }

            NozzleAssemblyComponent shellIntersectingComponent = FindShellIntersectingComponent();

            if (shellIntersectingComponent != null)
            {
                shellIntersectingComponent.Settings.IsShellIntersecting = true;

                Feature profileSketch = SWFeatureManager.GetFeatureByName(
                    shellIntersectingComponent.GetComponent(), "Profile Sketch");

                if (profileSketch != null)
                {
                    double diameter = profileSketch.Parameter("Diameter").GetValue3(
                        (int)swInConfigurationOpts_e.swAllConfiguration, null)[0];

                    shellIntersectingComponent.Settings.Diameter = diameter; // GetValue3 returns mm for this document
                }
            }

            // Add a cutout extrude feature (presumably to create space for the manhole assembly)
            AddCutOutExtrude(compartment, nozzle);

            // Update and save the active document and any associated attribute documents
            DocumentManager.UpdateAndSaveDocuments();
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
            Feature nozzleCutoutPlane = SWFeatureManager.GetFeatureByName(nozzleComp, "Nozzle Cutout Plane");

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
            string cutoutPlaneFullPath = $"{nozzleCutoutPlane.Name}@" +
                $"{ compartmentComp.Name2}@" +
                $"{ shellDoc.GetTitle()}/" +
                $"{ nozzleComp.Name2}@" +
                $"{ compartmentComp.Name2.Split('-')[0]}";

            bool planeSelected = shellDoc.Extension.SelectByID2(
                cutoutPlaneFullPath,
                "PLANE",
                0, 0, 0,
                false,
                0, null, 0);

            // Get access to the feature manager of the shell document.
            FeatureManager featureManager = shellDoc.FeatureManager;

            // Create a new reference plane coincident with the nozzle cutout plane and perpendicular to the nozzle axis.
            Feature cutOutPlane = (Feature)featureManager.InsertRefPlane(
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident,
                0,
                0,
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

            // 3. Extract Cutout Radius:
            // Diameter was already saved and flagged in AddNozzleAssembly — read it directly from settings.
            NozzleAssemblyComponent shellIntersectingComponent = nozzle._nozzleSettings.NozzleAssemblyComponents
                .FirstOrDefault(c => c.Settings.IsShellIntersecting);

            double radius = (shellIntersectingComponent.Settings.Diameter + CUT_CLEARANCE_MM) / 2.0 / 1000.0;

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
            Feature cutOutPlane = GetCutOutPlane();
            bool cutPlaneSelected = shellDoc.Extension.SelectByID2(
                          cutOutPlane.Name,
                          "PLANE",
                          0, 0, 0,
                          false,
                          0, null, 0);

            // Start a new sketch on the selected cutout plane.
            sketchManager.InsertSketch(true);

            // Enter sketch editing mode.
            shellDoc.EditSketch();

            // Create a circle on the sketch with the extracted radius, centered at a default position (0, 0, 0).
            SketchSegment circleSeg =  sketchManager.CreateCircleByRadius(0.1, 0, 0, radius);

            //Select the manhole axis and the center point of the circle for constraint application.
            // Get the center point of the circle and select it (append to keep axis selected)
            bool pointSelected = shellDoc.Extension.SelectByID2(
                           "Point2",
                           "SKETCHPOINT",
                           0, 0, 0,
                           false,
                           0, null, 0);

            bool axisSelected = shellDoc.Extension.SelectByID2(
                            $"{nozzleAxis.Name}@{compartmentComp.Name2}@{shellDoc.GetTitle()}/{nozzleComp.Name2}@{compartmentComp.Name2.Split('-')[0]}",
                            "AXIS",
                            0, 0, 0,
                            true,
                            0, null, 0);


            // Add a coincident constraint to align the circle's center with the manhole axis.
            shellDoc.SketchAddConstraints("sgCOINCIDENT");


            // Select the circle and add a driving diameter dimension (D1), then set it to the required value
            bool circleSelected = circleSeg.Select4(false, null);
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
            solidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            DisplayDimension displayDim = (DisplayDimension)shellDoc.AddDimension2(0.1, 0.1, 0);
            solidWorksApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, true);
            if (displayDim != null)
            {
                Dimension dim = displayDim.GetDimension2(0);
                dim.SetSystemValue3(
                    radius * 2,
                    (int)swSetValueInConfiguration_e.swSetValue_InThisConfiguration,
                    null);
            }

            // Get active sketch
            Sketch sketch = sketchManager.ActiveSketch;
            Feature sketchFeature = (Feature)sketch;
            sketchFeature.Name = $"{compartmentComp.Name2.Split('-')[0]} {nozzleComp.Name2.Split('-')[0]} Cut out sketch";
            string sketchName = sketchFeature.Name;

            //Start a new sketch on the selected cutout plane.
           sketchManager.InsertSketch(true);

            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.ActivateDocument();

            return sketchName;
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

            // Select the sketch that was created in AddCutOutSketch — FeatureCut4 requires an active sketch selection
            bool status = shellDoc.Extension.SelectByID2(sketchName, "SKETCH", 0, 0, 0, false, 0, null, 0);

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
                true,
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

            cutExtrude.Name = sketchName.Substring(0, sketchName.LastIndexOf(' ')) + " extrude";

            // Store a reference to the newly created cut-extrude feature in the manhole settings for later use.
            nozzle._nozzleSettings.PIDCutExtrude = shellDoc.Extension.GetPersistReference3(cutExtrude);

            
            // Update the SolidWorks documents to reflect the changes and save them.
            DocumentManager.UpdateAndSaveDocuments();

            // Replicate the macro sequence that makes the cut visible in the UI after the first nozzle.
            RefreshCutDisplay(cutExtrude);

        }

        /// <summary>
        /// Forces the Tank Site Assembly to rebuild and reset its editing context so that
        /// geometry changes (new cut, moved cut after offset change, etc.) appear correctly in the UI.
        /// Replicates the VBA macro: SelectByID2 → ClearSelection2 → AssemblyPartToggle → EditAssembly.
        /// </summary>
        public void RefreshCutDisplay(Feature cutExtrude)
        {
            // Select the cut body feature, then clear — this primes SolidWorks for the toggle.
            cutExtrude.Select2(false, 0);
            RefreshCutDisplay();
        }

        public void RefreshCutDisplay()
        {
            // Activate the Tank Site Assembly without triggering a rebuild on activation.
            ModelDoc2 tankSiteDoc = SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc;
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(
                tankSiteDoc.GetTitle(), true, (int)swRebuildOnActivation_e.swDontRebuildActiveDoc, 0);

            // Clear any active selections and disable contour selection mode,
            // matching the state expected by AssemblyPartToggle.
            tankSiteDoc.ClearSelection2(true);
            ((SelectionMgr)tankSiteDoc.SelectionManager).EnableContourSelection = false;

            // AssemblyPartToggle exits any open component-editing context and returns
            // focus to the top-level assembly. Without this call the cut feature remains
            // "owned" by the component context and does not render in the correct position.
            ((AssemblyDoc)tankSiteDoc).AssemblyPartToggle();

            // Re-enter the standard assembly editing state so subsequent operations work normally.
            ((AssemblyDoc)tankSiteDoc).EditAssembly();

            // Final clear to leave the document in a clean selection state.
            tankSiteDoc.ClearSelection2(true);
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

            SaveAndCloseDocument();
        }


        /// <summary>
        /// COMPARTMENT DOCUMENT MUST BE OPENED.
        /// Calculates the shortest distance (in meters) between the mid point (on the tank's center axis)
        /// and the internal point (on the inner wall of the cylindrical shell).
        /// Activates and closes the nozzle document internally.
        /// </summary>
        /// <returns>The distance in meters between the center axis and the inner wall.</returns>
        public double GetDistanceFromCenterAxisToInnerWall()
        {
            // Activate the nozzle document to access its features
            ModelDoc2 nozzleModelDoc = ActivateDocument();

            // Get the point where the tank's centerline meets the external wall
            Feature midPoint = GetMidPoint();

            // Get the top reference plane of the nozzle (cut plane)
            Feature innerPoint = GetInternalPoint();

            return GetDistanceBetweenTwoFeatures(nozzleModelDoc, midPoint, innerPoint);
        }


        /// <summary>
        /// COMPARTMENT DOCUMENT MUST BE OPENED.
        /// Calculates the shortest distance (in meters) between the point where the tank's
        /// centerline intersects the external wall and the nozzle's top reference plane (cut plane).
        /// Activates and closes the nozzle document internally.
        /// </summary>
        /// <returns>The distance in meters between the centerline-wall intersection and the cut plane.</returns>
        public double GetDistanceFromCenterlineWallIntersectionPointToNozzleTopPlane()
        {
            // Activate the nozzle document to access its features
            ModelDoc2 nozzleModelDoc = ActivateDocument();

            // Get the point where the tank's centerline meets the external wall
            Feature centerLineWallIntersection = GetCenterlineWallIntersection();

            // Get the top reference plane of the nozzle (cut plane)
            Feature plane = GetCutPlane();

            return GetDistanceBetweenTwoFeatures(nozzleModelDoc, centerLineWallIntersection, plane);
        }

        /// <summary>
        /// COMPARTMENT DOCUMENT MUST BE OPENED.
        /// Calculates the shortest distance (in meters) between the nozzle's internal point
        /// (where the nozzle meets the inner wall of the cylindrical shell) and the nozzle's
        /// top reference plane (cut plane).
        /// Activates and closes the nozzle document internally.
        /// </summary>
        /// <returns>The distance in meters between the internal point and the cut plane.</returns>
        public double GetDistanceFromNozzleInternalPointToNozzleTopPlane()
        {
            // Activate the nozzle document to access its features
            ModelDoc2 nozzleModelDoc = ActivateDocument();

            // Get the point where the tank's centerline meets the external wall
            Feature nozzleInternalPoint = GetInternalPoint();

            // Get the top reference plane of the nozzle (cut plane)
            Feature plane = GetCutPlane();

            return GetDistanceBetweenTwoFeatures(nozzleModelDoc, nozzleInternalPoint, plane);
        }

        /// <summary>
        /// Calculates the closest distance between two SolidWorks features using ModelDoc2.ClosestDistance.
        /// Always closes the nozzle document via the finally block, even if an exception occurs.
        /// The caller is responsible for activating the nozzle document and retrieving the features beforehand.
        /// </summary>
        /// <param name="nozzleModelDoc">The nozzle model document (must already be activated).</param>
        /// <param name="feature1">The first feature (point, plane, axis, etc.).</param>
        /// <param name="feature2">The second feature to measure distance to.</param>
        /// <returns>The shortest distance in meters between the two features.</returns>
        private double GetDistanceBetweenTwoFeatures(ModelDoc2 nozzleModelDoc, Feature feature1, Feature feature2)
        {
            try
            {
                // Calculate the closest distance between the intersection point and the plane
                double distance = nozzleModelDoc.ClosestDistance(
                   feature1,
                   feature2,
                   out object Point1,
                   out object Point2
                );

                return distance;
            }

            finally
            {
                CloseDocument();
            }
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

            SaveAndCloseDocument();
        }

        /// <summary>
        /// Rotates manhole according sketch circle.
        /// </summary>
        /// <param name="angleInDegrees"></param>
        public void SetRotationAngle(double angleInDegrees)
        {
            // Normalize angle to 0-360 degrees (using modulo operator)
            angleInDegrees = (angleInDegrees % 360 + 360) % 360;

            try
            {
                //Set angle
                Feature rotationMate = GetRotationMate();
                MateManager.ChangeMateAngle(rotationMate, angleInDegrees);
                _nozzleSettings.PIDRotationMate =
                    SolidWorksDocumentProvider.GetActiveDoc().Extension.GetPersistReference3(rotationMate);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error setting rotation angle: {ex.Message}");
            }

            SaveAndCloseDocument();
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
        public ModelDoc2 ActivateDocument()
        {
            // Get compartment doc
            ModelDoc2 nozzleModelDoc = GetComponent().GetModelDoc2();

            // Activate compartment doc
            _currentlyActiveNozzleDoc = SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(
                nozzleModelDoc.GetTitle() + ".sldasm", 
                true, 
                0, 
                0);

            return nozzleModelDoc;
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

        /// <summary>
        /// Saves the nozzle document then closes it.
        /// Use this instead of CloseDocument() after any modification so changes are not lost.
        /// </summary>
        public void SaveAndCloseDocument()
        {
            if (_currentlyActiveNozzleDoc == null) return;

            _currentlyActiveNozzleDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            CloseDocument();
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

        public Feature GetTopPoint() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDTopPoint,
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

        public Feature GetRotationMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDRotationMate,
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

        public Feature GetCenterlineWallIntersection() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                     _nozzleSettings.PIDCenterlineWallIntersection,
                     out int error);

        /// <summary>
        /// Finds the nozzle assembly component that intersects the shell by checking which component's
        /// mating and free planes straddle the external point (the point where the nozzle meets the outer
        /// shell wall). Works regardless of how the planes are oriented per component type.
        /// Nozzle document must be active before calling this method.
        /// </summary>
        /// <returns>The component that intersects the shell, or null if none found</returns>
        /// <summary>
        /// Positions the nozzle assembly along the nozzle axis so that the distance from the
        /// chosen top reference point to the first (topmost) component's outermost plane equals
        /// <paramref name="targetDistanceMeters"/>.
        ///
        /// Strategy – delta-based (no layout assumptions):
        ///
        ///   1. Measure the CURRENT distance from refPoint to the outermost plane of the first
        ///      component: currentDist.
        ///
        ///   2. The required shift equals (targetDistance - currentDist).
        ///      Apply that shift by adjusting the topPlaneMate distance by the same delta:
        ///
        ///          newMateDistance = currentMateDistance + (targetDistance - currentDist)
        ///
        /// This avoids any assumption about:
        ///   • which plane (mating or free) of the first component is mated with the Right Plane, and
        ///   • whether that mated plane is the outermost one.
        ///
        /// "Outermost" plane = whichever of matingPlane / freePlane has the GREATER signed
        /// distance from InsidePoint (always at the deepest inner end, so neither plane of the
        /// first component can fall on the wrong side of it — unlike refPoint, which can lie
        /// between the two planes when the free plane extends below refPoint).
        ///
        /// Nozzle document must be active before calling this method.
        /// </summary>
        public void SetTopReferenceDistance(NozzleTopReferenceType refType, double targetDistanceMeters)
        {
            if (_nozzleSettings.NozzleAssemblyComponents == null || _nozzleSettings.NozzleAssemblyComponents.Count == 0)
            {
                MessageBox.Show("No nozzle assembly components found.", "SetTopReferenceDistance");
                return;
            }

            // ── Step 1: resolve the reference point feature ──────────────────────────────────
            // TankCenterline → TopPoint;  NozzleCenterline → ExternalPoint
            Feature refPoint = refType == NozzleTopReferenceType.TankCenterline
                ? GetTopPoint()
                : GetExternalPoint();

            if (refPoint == null)
            {
                MessageBox.Show(
                    $"Could not retrieve reference point for {refType}.", "SetTopReferenceDistance");
                return;
            }

            // ── Step 2: outermost plane of the first (topmost) component ─────────────────────
            // Components are sorted top → bottom, so index 0 is the topmost.
            // "Outermost" = the plane closest to the reference point, regardless of whether it is
            // the mating plane or the free plane, and regardless of which one is mated with the
            // Right Plane.
            NozzleAssemblyComponent firstComponent = _nozzleSettings.NozzleAssemblyComponents[0];

            Feature matingPlane = firstComponent.GetMatingPlane();
            Feature freePlane   = firstComponent.GetFreePlane();
            Feature insidePoint = GetInsidePoint();

            if (matingPlane == null || freePlane == null || insidePoint == null)
            {
                MessageBox.Show("Could not retrieve mating/free plane or inside point.", "SetTopReferenceDistance");
                return;
            }

            // Use InsidePoint (deepest inner end) as reference to determine which plane is outermost.
            // InsidePoint is always on the opposite end from both planes of the first component,
            // so neither plane can be on the wrong side of it — unlike refPoint, which can lie
            // between the two planes (Case 2: free plane below refPoint).
            // The plane with the GREATER signed distance from InsidePoint is the outermost.
            double distMating = GetSignedDistanceAlongAxis(matingPlane, insidePoint, firstComponent);
            double distFree   = GetSignedDistanceAlongAxis(freePlane,   insidePoint, firstComponent);

            Feature outermostPlane = distMating >= distFree ? matingPlane : freePlane;
            double  currentDist    = _currentlyActiveNozzleDoc.ClosestDistance(refPoint, outermostPlane, out _, out _);

            // ── Step 3: current topPlaneMate distance ────────────────────────────────────────
            Feature topPlaneMate = GetTopPlaneMate();
            if (topPlaneMate == null)
            {
                MessageBox.Show("Top plane mate not found.", "SetTopReferenceDistance");
                return;
            }

            DistanceMateFeatureData mateData = topPlaneMate.GetDefinition() as DistanceMateFeatureData;
            if (mateData == null)
            {
                MessageBox.Show("Top plane mate is not a distance mate.", "SetTopReferenceDistance");
                return;
            }

            double currentMateDistance = mateData.Distance;

            // ── Step 4: delta and new mate distance ───────────────────────────────────────────
            //
            //   newMateDistance = currentMateDistance + (targetDistance - currentDist)
            //
            // Reasoning: shifting the assembly inward by Δ increases both currentDist and the
            // mate distance by exactly Δ.  So to make currentDist reach targetDistance we need
            // to increase the mate distance by (targetDistance - currentDist).
            double delta          = targetDistanceMeters - currentDist;
            double newMateDistance = currentMateDistance + delta;

            if (newMateDistance < 0)
            {
                MessageBox.Show(
                    $"Computed mate distance ({newMateDistance * 1000:F2} mm) is negative. " +
                    "The target distance is too small for the current nozzle geometry.",
                    "SetTopReferenceDistance");
                return;
            }

            // ── Step 5: apply ─────────────────────────────────────────────────────────────────
            MateManager.ChangeDistance(topPlaneMate, newMateDistance);

            // Keep the stored PID up to date
            _nozzleSettings.PIDTopPlaneMate =
                _currentlyActiveNozzleDoc.Extension.GetPersistReference3(topPlaneMate);
        }

        /// <summary>
        /// Adjusts the length of the last (bottommost) adjustable component so that the distance
        /// from its free plane to the selected reference point equals targetDistanceMeters.
        /// Nozzle document must be active before calling this method.
        /// </summary>
        public void SetAdjustableComponentLength(NozzleBottomReferencePoint refPoint, double targetDistanceMeters, bool isLongNozzle)
        {
            NozzleAssemblyComponent adjustableComponent = _nozzleSettings.NozzleAssemblyComponents
                .LastOrDefault(c => c.Settings.HasAdjustableLength);

            if (adjustableComponent == null)
            {
                string componentInfo = string.Join(", ", _nozzleSettings.NozzleAssemblyComponents
                    .Select(c => $"{c.Settings.ComponentType}(HasAdjustableLength={c.Settings.HasAdjustableLength})"));
                MessageBox.Show($"No adjustable component found.\nComponents: {componentInfo}", "SetAdjustableComponentLength");
                return;
            }

            Feature matingPlane = adjustableComponent.GetMatingPlane();
            if (matingPlane == null)
            {
                MessageBox.Show("Mating plane is null for adjustable component.", "SetAdjustableComponentLength");
                return;
            }

            Feature refFeature;
            switch (refPoint)
            {
                case NozzleBottomReferencePoint.Top:    refFeature = GetInternalPoint(); break;
                case NozzleBottomReferencePoint.Middle: refFeature = GetMidPoint();      break;
                case NozzleBottomReferencePoint.Bottom: refFeature = GetInsidePoint();   break;
                default: return;
            }

            if (refFeature == null)
            {
                MessageBox.Show($"Reference point feature is null for {refPoint}.", "SetAdjustableComponentLength");
                return;
            }

            double signedDist = GetSignedDistanceAlongAxis(matingPlane, refFeature, adjustableComponent);

            // +1: free plane must be further outward than ref → D1 = signedDist + target
            // -1: free plane must be closer (shallower) than ref → D1 = signedDist - target
            bool freeIsDeeper = refPoint == NozzleBottomReferencePoint.Top
                             || (refPoint == NozzleBottomReferencePoint.Middle && isLongNozzle);

            double newD1 = freeIsDeeper
                ? signedDist + targetDistanceMeters
                : signedDist - targetDistanceMeters;

            if (newD1 <= 0)
            {
                return;
            }

            adjustableComponent.ChangeLength(newD1);
        }

        /// <summary>
        /// Returns the signed distance from the reference point to the mating plane,
        /// measured along the plane's normal (which is the nozzle axis direction).
        /// Positive = mating plane is on the outward side of the reference point.
        /// </summary>
        private double GetSignedDistanceAlongAxis(Feature matingPlaneFeature, Feature refPointFeature, NozzleAssemblyComponent component)
        {
            SolidWorks.Interop.sldworks.MathUtility mathUtil =
                (SolidWorks.Interop.sldworks.MathUtility)SolidWorksDocumentProvider._solidWorksApplication.GetMathUtility();

            Component2 comp = component.GetComponent();
            if (comp == null) return 0;

            RefPlane matingRefPlane = (RefPlane)matingPlaneFeature.GetSpecificFeature2();
            if (matingRefPlane == null) return 0;

            MathTransform compTransform = comp.Transform2;

            // Plane normal (0,0,1) in component model space → nozzle assembly space
            SolidWorks.Interop.sldworks.MathVector normal =
                (SolidWorks.Interop.sldworks.MathVector)mathUtil.CreateVector(new double[] { 0, 0, 1 });
            normal = (SolidWorks.Interop.sldworks.MathVector)normal.MultiplyTransform(matingRefPlane.Transform);
            normal = (SolidWorks.Interop.sldworks.MathVector)normal.MultiplyTransform(compTransform);
            double[] n = (double[])normal.ArrayData;

            // Plane origin (0,0,0) in component model space → nozzle assembly space
            SolidWorks.Interop.sldworks.MathPoint planeOrigin =
                (SolidWorks.Interop.sldworks.MathPoint)mathUtil.CreatePoint(new double[] { 0, 0, 0 });
            planeOrigin = (SolidWorks.Interop.sldworks.MathPoint)planeOrigin.MultiplyTransform(matingRefPlane.Transform);
            planeOrigin = (SolidWorks.Interop.sldworks.MathPoint)planeOrigin.MultiplyTransform(compTransform);
            double[] o = (double[])planeOrigin.ArrayData;

            // Reference point position in nozzle assembly space
            RefPoint refPt = (RefPoint)refPointFeature.GetSpecificFeature2();
            if (refPt == null) return 0;
            double[] r = (double[])refPt.GetRefPoint().ArrayData;

            // dot(matingPlaneOrigin - refPos, planeNormal)
            // Positive when the mating plane is further outward than the reference point
            return (o[0] - r[0]) * n[0]
                 + (o[1] - r[1]) * n[1]
                 + (o[2] - r[2]) * n[2];
        }

        /// <summary>
        /// Returns the nozzle assembly components sorted from top to bottom
        /// distance from the "Inside point". The component whose planes are furthest from the
        /// Inside point is considered the topmost component.
        /// Nozzle document must be active before calling this method.
        /// </summary>
        public List<NozzleAssemblyComponent> GetComponentsSortedTopToBottom()
        {
            if (_nozzleSettings.NozzleAssemblyComponents == null || _nozzleSettings.NozzleAssemblyComponents.Count == 0)
                return new List<NozzleAssemblyComponent>();

            Feature insidePointFeature = GetInsidePoint();
            if (insidePointFeature == null)
                return new List<NozzleAssemblyComponent>(_nozzleSettings.NozzleAssemblyComponents);

            return _nozzleSettings.NozzleAssemblyComponents
                .OrderByDescending(c => GetComponentAverageDistanceToInsidePoint(c, insidePointFeature))
                .ToList();
        }

        /// <summary>
        /// Returns the signed axial distance from the Inside point to the component's mating plane.
        /// Uses the same transform-based approach as GetSignedDistanceAlongAxis for consistency.
        /// </summary>
        private double GetComponentAverageDistanceToInsidePoint(NozzleAssemblyComponent component, Feature insidePointFeature)
        {
            try
            {
                Feature matingPlaneFeature = component.GetMatingPlane();
                if (matingPlaneFeature == null) return 0;

                RefPoint insideRefPoint = (RefPoint)insidePointFeature.GetSpecificFeature2();
                if (insideRefPoint == null) return 0;

                double[] insidePos = (double[])insideRefPoint.GetRefPoint().ArrayData;

                // Reuse the same signed distance logic: project (matingOrigin - insidePos) onto the plane normal
                // in nozzle assembly space. This gives the axial distance, which is reliable for sorting.
                return GetSignedDistanceAlongAxis(matingPlaneFeature, insidePointFeature, component);
            }
            catch
            {
                return 0;
            }
        }

        public NozzleAssemblyComponent FindShellIntersectingComponent()
        {
            if (_nozzleSettings.NozzleAssemblyComponents == null || _nozzleSettings.NozzleAssemblyComponents.Count == 0)
                return null;

            // Get the external point feature and its XYZ coordinates
            Feature externalPointFeature = GetExternalPoint();
            if (externalPointFeature == null)
                return null;

            RefPoint externalRefPoint = (RefPoint)externalPointFeature.GetSpecificFeature2();
            if (externalRefPoint == null)
                return null;

            double[] externalCoords = (double[])externalRefPoint.GetRefPoint().ArrayData;

            // Check each component — return the first one whose mating and free planes straddle the external point
            foreach (var component in _nozzleSettings.NozzleAssemblyComponents)
            {
                if (component.IsShellIntersecting(externalCoords))
                    return component;
            }

            return null;
        }

        public void ChangeCutDiameterOfTankBodyEnvelope()
        {
            NozzleAssemblyComponent shellComp = _nozzleSettings.NozzleAssemblyComponents
                .FirstOrDefault(c => c.Settings.IsShellIntersecting);

            double diameterMm = shellComp?.Settings.Diameter + CUT_CLEARANCE_MM ?? 0;

            ModelDoc2 manholeDoc = ActivateDocument();

            // The TankBody_Envelope feature only exists in the "WithEnvelope" configuration.
            manholeDoc.ShowConfiguration2("WithEnvelope");

            string envelopeName = $"TankBody_Envelope^{manholeDoc.GetTitle()}";

            int errors = 0;
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc2(envelopeName, false, ref errors);
            ModelDoc2 envelopeDoc = SolidWorksDocumentProvider.GetActiveDoc();

            Feature cutOutSketch = SWFeatureManager.GetFeatureByName(envelopeDoc, "Cut out sketch");
            Dimension diameterDimension = cutOutSketch.Parameter("Diameter");
            diameterDimension.SetValue3(diameterMm, (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting, null);

            envelopeDoc.EditRebuild3();
            envelopeDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(envelopeDoc.GetTitle());

            manholeDoc.ShowConfiguration2("Default");

            manholeDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            CloseDocument();
        }

        public void ChangeTankBodyEnvelopeDimensions(
            double leftSideLength,
            double rightSideLength,
            double widthUp,
            double widthDown)
        {
            ModelDoc2 manholeDoc = ActivateDocument();

            // The TankBody_Envelope feature only exists in the "WithEnvelope" configuration.
            manholeDoc.ShowConfiguration2("WithEnvelope");
           
            string envelopeName = $"TankBody_Envelope^{manholeDoc.GetTitle()}";

            int errors = 0;
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc2(envelopeName, false, ref errors);
            ModelDoc2 envelopeDoc = SolidWorksDocumentProvider.GetActiveDoc();
            Feature tankBodySketch = SWFeatureManager.GetFeatureByName(envelopeDoc, "Tank body sketch");

            if (leftSideLength != null && leftSideLength > 0)
            {
                Dimension leftDimension = tankBodySketch.Parameter("LengthToLeft");
                leftDimension.SetValue3(leftSideLength, (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting, null);
            }
            if (rightSideLength != null && rightSideLength > 0)
            {
                Dimension leftrightDimension = tankBodySketch.Parameter("LengthToRight");
                leftrightDimension.SetValue3(rightSideLength, (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting, null);
            }

            Feature bossExtrude = SWFeatureManager.GetFeatureByName(envelopeDoc, "Boss-Extrude");
            ExtrudeFeatureData2 extrudeData = (ExtrudeFeatureData2)bossExtrude.GetDefinition();

            extrudeData.AccessSelections(envelopeDoc, null);

            if (widthUp != null && widthUp > 0)
            {
                extrudeData.SetDepth(true, widthUp / 1000);
            }
            if (widthDown != null && widthDown > 0)
            {
                extrudeData.SetDepth(false, widthDown / 1000);
            }
            bossExtrude.ModifyDefinition(extrudeData, envelopeDoc, null);

            envelopeDoc.EditRebuild3();
            envelopeDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(envelopeDoc.GetTitle());

            manholeDoc.ShowConfiguration2("Default");

            manholeDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);

            CloseDocument();
        }
    }
}

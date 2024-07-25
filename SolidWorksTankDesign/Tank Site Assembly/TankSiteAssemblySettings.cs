using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class TankSiteAssemblySettings
    {
        //--------------------------- Tank Site Assembly PIDs ---------------------------

        [JsonProperty("PIDCenterAxis")]
        public byte[] PIDCenterAxis { get; private set; }

        [JsonProperty("PIDWorkshopAssembly")]
        public byte[] PIDTankWorkshopAssembly { get; private set; }

        [JsonProperty("PIDAxisMate")]
        public byte[] PIDAxisMate { get; private set; }

        [JsonProperty("PIDTankAssembly")]
        public byte[] PIDTankAssembly { get; private set; }

        [JsonProperty("PIDShellAssembly")]
        public byte[] PIDShellAssembly { get; private set; }

        [JsonProperty("PIDDishedEndsAssembly")]
        public byte[] PIDDishedEndsAssembly { get; private set; }

        [JsonProperty("PIDCylindricalShellsAssembly")]
        public byte[] PIDCylindricalShellsAssembly { get; private set; }

        [JsonProperty("PIDCompartmentsAssembly")]
        public byte[] PIDCompartmentsAssembly { get; private set; }

        public TankSiteAssemblySettings() { }

        /// <summary>
        /// Collects and stores persistent reference IDs (PIDs) of key components and features within a tank site assembly model
        /// 
        /// Important Assumptions:
        /// - Center axis name - "Center axis".
        /// - Center axis mate name must be "Tank Workshop Assembly - Center axis".
        ///  -Attribute name - "MainEntities", attribute parameter name - "MainEntities".
        ///  -Order of subassemblies in Shell assembly: 1. Assembly of dished ends, 2. assembly of cylindrical shells, 3. assembly of compartments
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void AddTankSiteAssemblyPersistentReferenceIds(ModelDoc2 tankSiteModelDoc)
        {
            ModelDocExtension tankSiteModelDocExt = tankSiteModelDoc.Extension;
            SelectionMgr selectionMgr = (SelectionMgr)tankSiteModelDoc.SelectionManager;

            // Get components
            tankSiteModelDoc.Extension.SelectByID2(
                    "Center axis",
                    "AXIS",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Feature centerAxis = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly - Center axis",
                    "MATE",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Feature centerAxisMate = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 tankWorkshopAssemblyAsComponent = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly/Tank-1@Tank Workshop Assembly",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 tankAssemblyAsComponent = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly/Tank-1@Tank Workshop Assembly/Shell-1@Tank",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 shellAssemblyAsComponent = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly/Tank-1@Tank Workshop Assembly/Shell-1@Tank/Assembly of Dished ends-1@Shell",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 assemblyOfDishedEndsAsComponent = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly/Tank-1@Tank Workshop Assembly/Shell-1@Tank/Assembly of Cylindrical Shells-1@Shell",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 assemblyOfCylindricalShellsAsComponent = selectionMgr.GetSelectedObject6(1, -1);

            tankSiteModelDoc.Extension.SelectByID2(
                    "Tank Workshop Assembly-1@Tank Site Assembly/Tank-1@Tank Workshop Assembly/Shell-1@Tank/Compartments-1@Shell",
                    "COMPONENT",
                    0, 0, 0,
                    false,
                    0, null, 0);
            Component2 assemblyOfCompartmentsAsComponent = selectionMgr.GetSelectedObject6(1, -1);
            

            // Safety check: Ensure that components were found
            if (tankWorkshopAssemblyAsComponent == null)
            {
                throw new InvalidOperationException("Tank workshop assembly component not found in the tank site assembly.");
            }
            if (tankAssemblyAsComponent == null)
            {
                throw new InvalidOperationException("Tank assembly component was not found in Tank Workshop Assembly.");
            }
            if (shellAssemblyAsComponent == null)
            {
                throw new InvalidOperationException("Shell assembly component was not found in Tank Assembly.");
            }

            try
            {
                //------------------------- Tank Site Assembly PIDs ----------------------------------------

                // Add persistent reference IDs
                PIDCenterAxis = tankSiteModelDocExt.GetPersistReference3(centerAxis);
                PIDTankWorkshopAssembly = tankSiteModelDocExt.GetPersistReference3(tankWorkshopAssemblyAsComponent);
                PIDAxisMate = tankSiteModelDocExt.GetPersistReference3(centerAxisMate);
                PIDTankAssembly = tankSiteModelDocExt.GetPersistReference3(tankAssemblyAsComponent);
                PIDShellAssembly = tankSiteModelDocExt.GetPersistReference3(shellAssemblyAsComponent);
                PIDDishedEndsAssembly = tankSiteModelDocExt.GetPersistReference3(assemblyOfDishedEndsAsComponent);
                PIDCylindricalShellsAssembly = tankSiteModelDocExt.GetPersistReference3(assemblyOfCylindricalShellsAsComponent);
                PIDCompartmentsAssembly = tankSiteModelDocExt.GetPersistReference3(assemblyOfCompartmentsAsComponent);


                //----------------------- Disehd Ends PIDs ---------------------------------------------------


            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message, "The attribute could not be created.");
                return;
            }
        }

        /// <summary>
        /// Adds persistent reference IDs (PIDs) to the settings of the left and right dished ends components within an assembly of dished ends.
        /// </summary>
        /// <param name="tankSiteModelDoc">The main SolidWorks model document containing the tank site assembly.</param>
        /// <returns>An `AssemblyOfDishedEnds` object with the left dished end's PIDs populated, or null if an error occurs.</returns>
        /// <remarks>
        /// **Important Assumptions:**
        /// - The left dished end component is named "Left dished end-1".
        /// - The position plane is named "Left end plane".
        /// - The center axis is named "Center axis".
        /// - The center axis mate is named "Left dished end - Center axis".
        /// - The right plane mate is named "Left dished end - Right plane".
        /// - The front plane mate is named "Left dished end - Front plane".
        /// 
        /// - The right dished end component is named "Right dished end-2"
        /// - The position plane is named "Right end plane".
        /// - The center axis is named "Center axis".
        /// - The center axis mate is named "Right dished end - Center axis".
        /// - The right plane mate is named "Right dished end - Right plane".
        /// - The front plane mate is named "Right dished end - Front plane".
        /// 
        /// **Workflow:**
        /// 1. Retrieves the dished ends assembly component using its PID (`PIDDishedEndsAssembly`).
        /// 2. Locates the dished end components within the assembly.
        /// 3. Creates an `AssemblyOfDishedEnds` object.
        /// 4. Retrieves PIDs for relevant features/mates of the left and right dished ends.
        /// 5. Populates the `_dishedEndSettings` of the dished ends with the retrieved PIDs.
        /// 6. Returns the `AssemblyOfDishedEnds` object with initialized settings, or null if errors occur.
        /// </remarks>
        public AssemblyOfDishedEnds AddDishedEndsPIDs(ModelDoc2 tankSiteModelDoc)
        {
            // Retrieve the dished ends assembly component
            Component2 dishedEndsComponent = (Component2)tankSiteModelDoc.Extension.GetObjectByPersistReference3(PIDDishedEndsAssembly, out _);
            if (dishedEndsComponent == null)
            {
                MessageBox.Show("Assembly of dished ends was not found.");
                return null;
            }
            
            ModelDoc2 dishedEndsModelDoc = dishedEndsComponent.GetModelDoc2();

            using(var document = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, dishedEndsModelDoc))
            {
                SelectionMgr selectionMgr = (SelectionMgr)dishedEndsModelDoc.SelectionManager;

                // Create an AssemblyOfDishedEnds object to represent the assembly and store its settings.
                AssemblyOfDishedEnds assemblyOfDishedEnds = new AssemblyOfDishedEnds(dishedEndsModelDoc);

                GetLeftDishedEndEntitiesPIDs();

                GetRightDishedEndEntitiesPIDs();

                // Return the AssemblyOfDishedEnds object with populated PIDs for the left dished end.
                return assemblyOfDishedEnds;

                void GetLeftDishedEndEntitiesPIDs()
                {
                    // Get features and components
                    dishedEndsModelDoc.Extension.SelectByID2(
                       "Left end plane",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature positionPlane = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Left dished end-1@Assembly of Dished ends",
                        "COMPONENT",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Component2 leftDishedEndComponent = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Center axis@Left dished end-1@Assembly of Dished ends",
                        "AXIS",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxis = selectionMgr.GetSelectedObject6(1, -1);


                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Left dished end - Center axis",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxisMate = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                         "Left dished end - Right plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature rightPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                         "Left dished end - Front plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature frontPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    try
                    {
                        // Populate the _dishedEndSettings with the retrieved PIDs
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDPositionPlane = dishedEndsModelDoc.Extension.GetPersistReference3(positionPlane);
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDComponent = dishedEndsModelDoc.Extension.GetPersistReference3(leftDishedEndComponent);
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDCenterAxis = dishedEndsModelDoc.Extension.GetPersistReference3(centerAxis);
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDCenterAxisMate = dishedEndsModelDoc.Extension.GetPersistReference3(centerAxisMate);
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDRightPlaneMate = dishedEndsModelDoc.Extension.GetPersistReference3(rightPlaneMate);
                        assemblyOfDishedEnds.LeftDishedEnd._dishedEndSettings.PIDFrontPlaneMate = dishedEndsModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "The attribute could not be created.");
                        return;
                    }
                }

                void GetRightDishedEndEntitiesPIDs()
                {
                    // Get features and components
                    dishedEndsModelDoc.Extension.SelectByID2(
                       "Right end plane",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature positionPlane = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Right dished end-2@Assembly of Dished ends",
                        "COMPONENT",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Component2 rightDishedEndComponent = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Center axis@Right dished end-2@Assembly of Dished ends",
                        "AXIS",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxis = selectionMgr.GetSelectedObject6(1, -1);


                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Right dished end - Center axis",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxisMate = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                         "Right dished end - Right plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature rightPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    dishedEndsModelDoc.Extension.SelectByID2(
                        "Right dished end - Front plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature frontPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    try
                    {
                        // Populate the _dishedEndSettings with the retrieved PIDs
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDPositionPlane = dishedEndsModelDoc.Extension.GetPersistReference3(positionPlane);
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDComponent = dishedEndsModelDoc.Extension.GetPersistReference3(rightDishedEndComponent);
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDCenterAxis = dishedEndsModelDoc.Extension.GetPersistReference3(centerAxis);
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDCenterAxisMate = dishedEndsModelDoc.Extension.GetPersistReference3(centerAxisMate);
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDRightPlaneMate = dishedEndsModelDoc.Extension.GetPersistReference3(rightPlaneMate);
                        assemblyOfDishedEnds.RightDishedEnd._dishedEndSettings.PIDFrontPlaneMate = dishedEndsModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "The attribute could not be created.");
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Adds persistent reference IDs (PIDs) to the settings of cylindrical shell component within an assembly of cylindrical shells.
        /// </summary>
        /// <param name="tankSiteModelDoc">The main SolidWorks model document containing the tank site assembly.</param>
        /// <returns>An `AssemblyCylindricalShells` object with the first cylindrical shell's PIDs populated, or null if an error occurs.</returns>
        /// <remarks>
        /// **Important Assumptions:**
        /// - The first cylindrical shell component is named "Cylindrical shell-1".
        /// - Left end plane name "Left End Plane"
        /// - Right end plane name "Right End Plane"
        /// - The center axis is named "Center axis".
        /// - The center axis mate is named "Cylindrical shell 1 - Center axis".
        /// - The left end mate is named "Cylindrical shell 1 - Left plane".
        /// - The front plane mate is named "Cylindrical shell 1 - Front plane".
        /// - First cylindrical shells' Front Mate Angle must be 45 degrees.
        /// 
        /// **Workflow:**
        /// 1. Retrieves the cylindrical shells assembly component using its PID (`PIDCylindricalShellsAssembly`).
        /// 2. Locates the cylindrical shell component within the assembly.
        /// 3. Creates an `AssemblyOfCylindricalShells` object.
        /// 4. Retrieves PIDs for relevant features/mates of the first cylindrical shell.
        /// 5. Populates the `_cylindricalShellSettings` of the first cylindrical shell with the retrieved PIDs.
        /// 6. Returns the `AssemblyOfCylindricalShells` object with initialized settings, or null if errors occur.
        /// </remarks>
        public AssemblyOfCylindricalShells AddCylindricalShellsPIDs(ModelDoc2 tankSiteModelDoc)
        {
            // Retrieve the cylindrical shells assembly component
            Component2 cylindricalShellsComponent = (Component2)tankSiteModelDoc.Extension.GetObjectByPersistReference3(PIDCylindricalShellsAssembly, out _);
            if (cylindricalShellsComponent == null)
            {
                MessageBox.Show("Assembly of cylindrical shells was not found.");
                return null;
            }

            ModelDoc2 cylindricalShellsModelDoc = cylindricalShellsComponent.GetModelDoc2();

            using (var document = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, cylindricalShellsModelDoc))
            {
                SelectionMgr selectionMgr = (SelectionMgr)cylindricalShellsModelDoc.SelectionManager;

                // Create an AssemblyOfCylindricalShells object to represent the assembly and store its settings.
                AssemblyOfCylindricalShells assemblyOfCylindricalShells = new AssemblyOfCylindricalShells();

                // Create the object of the first cylindrical shell and add it to the cylindrical shells list
                CylindricalShell cylindricalShell = new CylindricalShell();
                assemblyOfCylindricalShells.CylindricalShells.Add(cylindricalShell);

                GetCylindricalShellEntitiesPIDs();

                return assemblyOfCylindricalShells;

                void GetCylindricalShellEntitiesPIDs()
                {
                    // Get features and components
                    cylindricalShellsModelDoc.Extension.SelectByID2(
                       "Left End Plane@Cylindrical shell-1@Assembly of Cylindrical Shells",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature leftEndPlane = selectionMgr.GetSelectedObject6(1, -1);

                    // Get features and components
                    cylindricalShellsModelDoc.Extension.SelectByID2(
                       "Right End Plane@Cylindrical shell-1@Assembly of Cylindrical Shells",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature rightEndPlane = selectionMgr.GetSelectedObject6(1, -1);
                    
                    // Get features and components
                    cylindricalShellsModelDoc.Extension.SelectByID2(
                       "Cylindrical shell-1@Assembly of Cylindrical Shells",
                       "COMPONENT",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Component2 cylindricalShellComponent = selectionMgr.GetSelectedObject6(1, -1);

                     cylindricalShellsModelDoc.Extension.SelectByID2(
                       "Center Axis@Cylindrical shell-1@Assembly of Cylindrical Shells",
                       "AXIS",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature centerAxis = selectionMgr.GetSelectedObject6(1, -1);

                    cylindricalShellsModelDoc.Extension.SelectByID2(
                        "Cylindrical shell 1 - Left plane",
                       "MATE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature leftEndMate = selectionMgr.GetSelectedObject6(1, -1);

                    cylindricalShellsModelDoc.Extension.SelectByID2(
                         "Cylindrical shell 1 - Front plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature frontPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    cylindricalShellsModelDoc.Extension.SelectByID2(
                        "Cylindrical shell 1 - Center axis",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxisMate = selectionMgr.GetSelectedObject6(1, -1);

                    try
                    {
                        // Populate the _dishedEndSettings with the retrieved PIDs
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDLeftEndPlane = cylindricalShellsModelDoc.Extension.GetPersistReference3(leftEndPlane);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDRightEndPlane = cylindricalShellsModelDoc.Extension.GetPersistReference3(rightEndPlane);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDComponent = cylindricalShellsModelDoc.Extension.GetPersistReference3(cylindricalShellComponent);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDCenterAxis = cylindricalShellsModelDoc.Extension.GetPersistReference3(centerAxis);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDLeftEndMate = cylindricalShellsModelDoc.Extension.GetPersistReference3(leftEndMate);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDFrontPlaneMate = cylindricalShellsModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                        assemblyOfCylindricalShells.CylindricalShells[0]._cylindricalShellSettings.PIDCenterAxisMate = cylindricalShellsModelDoc.Extension.GetPersistReference3(centerAxisMate);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "The attribute could not be created.");
                        return;
                    }

                    // Return the AssemblyOfDishedEnds object with populated PIDs for the left dished end.
                    
                }
            }
        }

        /// <summary>
        /// Initializes and populates persistent reference IDs (PIDs) for the compartment and its associated nozzle within the tank site assembly model.
        /// </summary>
        /// <param name="tankSiteModelDoc">The SolidWorks model document representing the main tank site assembly.</param>
        /// <returns>
        /// A `CompartmentsManager` object with PIDs for the first compartment and its nozzle initialized,
        /// or `null` if an error occurs during PID retrieval or component/feature access.
        /// </returns>
        /// <remarks>
        /// This method specifically targets the first compartment (named "Compartment A Manholes-1") and its first nozzle (named "Manhole 1") within the shell assembly.
        ///
        /// **Component and Feature Naming Assumptions:**
        /// * Shell Assembly: "Shell"
        /// * Compartment: "Compartment A Manholes-1"
        ///     * Left End Plane: "Left end plane"
        ///     * Right End Plane: "Right end plane"
        ///     * Center Axis: "Center Axis"
        ///     * Left End Mate: "Compartment A - Dished end position plane"
        ///     * Front Plane Mate: "Compartment A - Front plane"
        ///     * Center Axis Mate: "Compartment A - Center axis"
        /// * Nozzle: "Manhole 1"
        ///     * Position Plane: "Manhole 1 position plane"
        ///     * Nozzle Component: "M1 Manhole-1@Compartment A Manholes"
        ///     * Nozzle Center Axis: "Center Axis"
        ///     * Nozzle Points (Datum): "External point", "Internal point", "Inside point", "Mid point"
        ///     * Nozzle Right Reference Plane: "Nozzle Right Reference Plane"
        ///     * ... (Additional nozzle features: PLANE1, Sketch, Manhole DN600, etc.)
        ///
        /// **Error Handling:**
        /// * Displays a MessageBox if the shell assembly cannot be found.
        /// * Displays a MessageBox with exception details if a PID cannot be retrieved or an entity (feature, component, mate) cannot be accessed.
        ///
        /// **Workflow:**
        /// 1. Accesses the shell assembly component using its PID.
        /// 2. Creates `CompartmentsManager` and `Compartment` objects to store settings.
        /// 3. Iterates over compartment and nozzle features/components/mates.
        /// 4. Uses `SelectByID2` to select each entity by its name within the appropriate context (shell, compartment, or nozzle).
        /// 5. Retrieves the PID for the selected entity using `GetPersistReference3`.
        /// 6. Populates the corresponding PID property within the `_compartmentSettings` or `_nozzleSettings` objects.
        /// 7. Returns the populated `CompartmentsManager` object or null if errors occur.
        /// </remarks>
        public CompartmentsManager AddCompartmentsManagerPIDs(ModelDoc2 tankSiteModelDoc)
        {
            // Retrieve the shell assembly component
            Component2 shellComponent = (Component2)tankSiteModelDoc.Extension.GetObjectByPersistReference3(PIDShellAssembly, out _);
            if (shellComponent == null)
            {
                MessageBox.Show("Assembly of shell was not found.");
                return null;
            }

            // Get the SolidWorks model document associated with the shell assembly component.
            ModelDoc2 shellModelDoc = shellComponent.GetModelDoc2();

            // Use a SolidWorksDocumentWrapper for managing the shell model document.
            using (var document = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, shellModelDoc))
            {
                // Get the selection manager to interact with selections within the shell model.
                SelectionMgr selectionMgr = (SelectionMgr)shellModelDoc.SelectionManager;

                // Create a shell object to represent the assembly and store its settings.
                CompartmentsManager compartmentManager = new CompartmentsManager();

                // Create the object of the first compartment and add it to the comparments list in the shell
                Compartment compartment = new Compartment();
                compartmentManager.Compartments.Add(compartment);

                Component2 compartmentComponent;

                // Call the helper method to retrieve and store PIDs for compartment entities.
                GetCompartmentEntitiesPIDs();

                // Create a nozzle object representing the first nozzle associated with the compartment.
                Nozzle nozzle = new Nozzle();
                compartmentManager.Compartments[0].Nozzles.Add(nozzle); // Add the nozzle to the compartment's list.

                // Call the helper method to retrieve and store PIDs for nozzle entities.
                GetNozzleEntitiesPIDs();

                return compartmentManager;

                // Helper method to retrieve and store PIDs for features and components within the compartment.
                void GetCompartmentEntitiesPIDs()
                {
                    // Get features and components
                    shellModelDoc.Extension.SelectByID2(
                       "Compartment A Manholes-1@Shell",
                       "COMPONENT",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    compartmentComponent = selectionMgr.GetSelectedObject6(1, -1);

                    shellModelDoc.Extension.SelectByID2(
                        "Compartment A - Dished end position plane",
                       "MATE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature leftEndMate = selectionMgr.GetSelectedObject6(1, -1);

                    ModelDoc2 compartmentModelDoc = compartmentComponent.GetModelDoc2();
                    using (var compartmentDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, compartmentModelDoc))
                    {
                        // Get the selection manager to interact with selections within the shell model.
                        SelectionMgr selectionMgrAtCompartmentDoc = (SelectionMgr)compartmentModelDoc.SelectionManager;

                        compartmentModelDoc.Extension.SelectByID2(
                            "Left end plane",
                            "PLANE",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature leftEndPlane = selectionMgrAtCompartmentDoc.GetSelectedObject6(1, -1);

                        compartmentModelDoc.Extension.SelectByID2(
                           "Right end plane",
                           "PLANE",
                           0, 0, 0,
                           false,
                           0, null, 0);
                        Feature rightEndPlane = selectionMgrAtCompartmentDoc.GetSelectedObject6(1, -1); compartmentModelDoc.Extension.SelectByID2(
                            "Center Axis",
                            "AXIS",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        Feature centerAxis = selectionMgrAtCompartmentDoc.GetSelectedObject6(1, -1);

                        compartmentManager.Compartments[0]._compartmentSettings.PIDLeftEndPlane = compartmentModelDoc.Extension.GetPersistReference3(leftEndPlane);
                        compartmentManager.Compartments[0]._compartmentSettings.PIDRightEndPlane = compartmentModelDoc.Extension.GetPersistReference3(rightEndPlane);
                        compartmentManager.Compartments[0]._compartmentSettings.PIDCenterAxis = compartmentModelDoc.Extension.GetPersistReference3(centerAxis);
                    }

                    try
                    {
                        // Populate the _compartmentSettings with the retrieved PIDs
                        compartmentManager.Compartments[0]._compartmentSettings.PIDComponent = shellModelDoc.Extension.GetPersistReference3(compartmentComponent);
                        compartmentManager.Compartments[0]._compartmentSettings.PIDLeftEndMate = shellModelDoc.Extension.GetPersistReference3(leftEndMate);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "The attribute could not be created.");
                        return;
                    }

                }

                // Helper method to retrieve and store PIDs for features and components within the nozzle.
                void GetNozzleEntitiesPIDs()
                {
                    // Get the SolidWorks model document associated with the compartment's assembly component.
                    ModelDoc2 compartmentModelDoc = compartmentComponent.GetModelDoc2();

                    Component2 nozzle1;

                    // Use a SolidWorksDocumentWrapper for managing the compartment's model document.
                    using (var compartmentDocument = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, compartmentModelDoc))
                    {
                        // Get the selection manager to interact with selections within the compartment's model
                        SelectionMgr selectionMgrAtCompartment = (SelectionMgr)compartmentModelDoc.SelectionManager;

                        // Get features and components
                       compartmentModelDoc.Extension.SelectByID2(
                            "Manhole 1 position plane",
                            "PLANE",
                            0, 0, 0,
                            false,
                            0, null, 0);
                       Feature nozzle1PostionPlane = selectionMgrAtCompartment.GetSelectedObject6(1, -1);

                       compartmentModelDoc.Extension.SelectByID2(
                            "M1 Manhole-1@Compartment A Manholes",
                            "COMPONENT",
                            0, 0, 0,
                            false,
                            0, null, 0);
                        nozzle1 = selectionMgrAtCompartment.GetSelectedObject6(1, -1);

                        compartmentModelDoc.Extension.SelectByID2(
                           "M1 - Position plane",
                           "MATE",
                           0, 0, 0,
                           false,
                           0, null, 0);
                        Feature positionPlaneMate = selectionMgrAtCompartment.GetSelectedObject6(1, -1);

                        try
                        {
                            // Populate the _nozzleSettings with the retrieved PIDs for those entities that has to be reachable from compartment's document
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDPositionPlane = compartmentModelDoc.Extension.GetPersistReference3(nozzle1PostionPlane);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDComponent = compartmentModelDoc.Extension.GetPersistReference3(nozzle1);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDPositionPlaneMate = compartmentModelDoc.Extension.GetPersistReference3(positionPlaneMate);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "The attribute could not be created.");
                            return;
                        }
                    }

                    // Get the SolidWorks model document associated with the nozzle's assembly component.
                    ModelDoc2 nozzleModelDoc = nozzle1.GetModelDoc2();

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
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDCenterAxis = nozzleModelDoc.Extension.GetPersistReference3(nozzleCenterAxis);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDExternalPoint = nozzleModelDoc.Extension.GetPersistReference3(externalPoint);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDInternalPoint = nozzleModelDoc.Extension.GetPersistReference3(internalPoint);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDInsidePoint = nozzleModelDoc.Extension.GetPersistReference3(insidePoint);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDMidPoint = nozzleModelDoc.Extension.GetPersistReference3(midPoint);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDNozzleRightRefPlane = nozzleModelDoc.Extension.GetPersistReference3(nozzleRightRefPlane);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDCutPlane = nozzleModelDoc.Extension.GetPersistReference3(plane1);
                            compartmentManager.Compartments[0].Nozzles[0]._nozzleSettings.PIDSketch = nozzleModelDoc.Extension.GetPersistReference3(sketch);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "The attribute could not be created.");
                            return;
                        }
                    }
                }

            }
        }
    }
}

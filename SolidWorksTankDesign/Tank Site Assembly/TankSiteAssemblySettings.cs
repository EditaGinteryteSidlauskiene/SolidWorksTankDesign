using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
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
        public Shell AddShellPIDs(ModelDoc2 tankSiteModelDoc)
        {
            // Retrieve the cylindrical shells assembly component
            Component2 shellComponent = (Component2)tankSiteModelDoc.Extension.GetObjectByPersistReference3(PIDShellAssembly, out _);
            if (shellComponent == null)
            {
                MessageBox.Show("Assembly of shell was not found.");
                return null;
            }

            ModelDoc2 shellModelDoc = shellComponent.GetModelDoc2();

            using (var document = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, shellModelDoc))
            {
                SelectionMgr selectionMgr = (SelectionMgr)shellModelDoc.SelectionManager;

                // Create a shell object to represent the assembly and store its settings.
                Shell shell = new Shell();

                // Create the object of the first compartment and add it to the comparments list in the shell
                Compartment compartment = new Compartment();
                shell.Compartments.Add(compartment);

                GetCompartmentEntitiesPIDs();

                return shell;


                void GetCompartmentEntitiesPIDs()
                {
                    // Get features and components
                    shellModelDoc.Extension.SelectByID2(
                       "Left end plane@Compartment A Manholes-1@Shell",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature leftEndPlane = selectionMgr.GetSelectedObject6(1, -1);

                    // Get features and components
                    shellModelDoc.Extension.SelectByID2(
                       "Right end plane@Compartment A Manholes-1@Shell",
                       "PLANE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature rightEndPlane = selectionMgr.GetSelectedObject6(1, -1);

                    // Get features and components
                    shellModelDoc.Extension.SelectByID2(
                       "Compartment A Manholes-1@Shell",
                       "COMPONENT",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Component2 compartmentComponent = selectionMgr.GetSelectedObject6(1, -1);

                    shellModelDoc.Extension.SelectByID2(
                      "Center Axis@Compartment A Manholes-1@Shell",
                      "AXIS",
                      0, 0, 0,
                      false,
                      0, null, 0);
                    Feature centerAxis = selectionMgr.GetSelectedObject6(1, -1);

                    shellModelDoc.Extension.SelectByID2(
                        "Compartment A - Dished end position plane",
                       "MATE",
                       0, 0, 0,
                       false,
                       0, null, 0);
                    Feature leftEndMate = selectionMgr.GetSelectedObject6(1, -1);

                    shellModelDoc.Extension.SelectByID2(
                        "Compartment A - Front plane",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature frontPlaneMate = selectionMgr.GetSelectedObject6(1, -1);

                    shellModelDoc.Extension.SelectByID2(
                        "Compartment A - Center axis",
                        "MATE",
                        0, 0, 0,
                        false,
                        0, null, 0);
                    Feature centerAxisMate = selectionMgr.GetSelectedObject6(1, -1);

                    try
                    {
                        // Populate the _dishedEndSettings with the retrieved PIDs
                        shell.Compartments[0]._compartmentSettings.PIDLeftEndPlane = shellModelDoc.Extension.GetPersistReference3(leftEndPlane);
                        shell.Compartments[0]._compartmentSettings.PIDRightEndPlane = shellModelDoc.Extension.GetPersistReference3(rightEndPlane);
                        shell.Compartments[0]._compartmentSettings.PIDComponent = shellModelDoc.Extension.GetPersistReference3(compartmentComponent);
                        shell.Compartments[0]._compartmentSettings.PIDCenterAxis = shellModelDoc.Extension.GetPersistReference3(centerAxis);
                        shell.Compartments[0]._compartmentSettings.PIDLeftEndMate = shellModelDoc.Extension.GetPersistReference3(leftEndMate);
                        shell.Compartments[0]._compartmentSettings.PIDFrontPlaneMate = shellModelDoc.Extension.GetPersistReference3(frontPlaneMate);
                        shell.Compartments[0]._compartmentSettings.PIDCenterAxisMate = shellModelDoc.Extension.GetPersistReference3(centerAxisMate);
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
    }
}

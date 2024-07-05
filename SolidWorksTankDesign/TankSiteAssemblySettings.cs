using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
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
        public AssemblyOfDishedEnds AddDishedEndsPIDs(SldWorks solidWorksApplication, ModelDoc2 tankSiteModelDoc)
        {
            // Retrieve the dished ends assembly component
            Component2 dishedEndsComponent = (Component2)tankSiteModelDoc.Extension.GetObjectByPersistReference3(PIDDishedEndsAssembly, out _);
            if (dishedEndsComponent == null)
            {
                MessageBox.Show("Assembly of dished ends was not found.");
                return null;
            }
            
            ModelDoc2 dishedEndsModelDoc = dishedEndsComponent.GetModelDoc2();

            using(var document = new SolidWorksDocumentWrapper(solidWorksApplication, dishedEndsModelDoc))
            {
                SelectionMgr selectionMgr = (SelectionMgr)dishedEndsModelDoc.SelectionManager;

                dishedEndsModelDoc.Extension.SelectByID2(
                        "Left dished end-1@Assembly of Dished ends",
                        "COMPONENT",
                        0, 0, 0,
                        false,
                        0, null, 0);
                Component2 leftDishedEndComponent = selectionMgr.GetSelectedObject6(1, -1);

                dishedEndsModelDoc.Extension.SelectByID2(
                        "Right dished end-2@Assembly of Dished ends",
                        "COMPONENT",
                        0, 0, 0,
                        false,
                        0, null, 0);
                Component2 rightDishedEndComponent = selectionMgr.GetSelectedObject6(1, -1);


                List<Component2>innerDishedEndComponentList = new List<Component2>();

                // Create an AssemblyOfDishedEnds object to represent the assembly and store its settings.
                AssemblyOfDishedEnds assemblyOfDishedEnds = new AssemblyOfDishedEnds(solidWorksApplication, dishedEndsModelDoc, leftDishedEndComponent, rightDishedEndComponent);

                GetLeftDishedEndEntities();

                GetRightDishedEndEntities();

                // Return the AssemblyOfDishedEnds object with populated PIDs for the left dished end.
                return assemblyOfDishedEnds;

                void GetLeftDishedEndEntities()
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

                void GetRightDishedEndEntities()
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

    }
}

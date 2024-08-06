using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System.Windows.Forms;
using System;
using System.IO;
using System.Xml.Linq;
using AddinWithTaskpane;
using SolidWorks.Interop.swconst;
using System.ComponentModel;

namespace SolidWorksTankDesign
{
    internal class CylindricalShell
    {
        private const string CYLINDRICAL_SHELL_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\Shell Cylyndrical ø1600×6 L1000_2_638584524028007781.SLDPRT";
        private const string CYLINDRICAL_SHELL_COMPONENT_NAME = "Cylindrical shell";
        private const string LEFT_END_PLANE_NAME = "Left plane";
        private const string CENTER_AXIS_NAME = "Center axis";
        private const string FRONT_PLANE_NAME = "Front plane";

        public CylindricalShellSettings _cylindricalShellSettings;

        /// <summary>
        /// DO NOT DELETE IT!!! This constructor is needed for Json deserializer.
        /// </summary>
        public CylindricalShell() 
        {
            _cylindricalShellSettings = new CylindricalShellSettings();
        }

        /// <summary>
        /// This constructor is called when a new cylindrical shell is added.
        /// Creates and positions a cylindrical shell within a SolidWorks assembly.
        /// </summary>
        /// <param name="referenceCylindricalShell"></param>
        /// <param name="assemblyOfCylindricalShellsCenterAxis"></param>
        /// <param name="assemblyOfCylindricalShellsFrontPlane"></param>
        /// <param name="length"></param>
        /// <param name="diameter"></param>
        /// <param name="countNumber"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public CylindricalShell(
            CylindricalShell referenceCylindricalShell,
            Feature assemblyOfCylindricalShellsCenterAxis,
            Feature assemblyOfCylindricalShellsFrontPlane,
            double length,
            double diameter,
            int countNumber)
        {
            // 1. Input Validation
            if (assemblyOfCylindricalShellsCenterAxis == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsCenterAxis));

            if (assemblyOfCylindricalShellsFrontPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsFrontPlane));

            if (referenceCylindricalShell == null)
                throw new ArgumentNullException(nameof(referenceCylindricalShell));

            ModelDoc2 assemblyOfCylindricalShellsDoc = SolidWorksDocumentProvider.GetActiveDoc();
            if (assemblyOfCylindricalShellsDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            // 2. Add and Make Independent Cylindrical Shell Component
            Component2 cylindricalShell = ComponentManager.AddComponentPart(CYLINDRICAL_SHELL_PATH);
            
            ComponentManager.MakeComponentIndependent(cylindricalShell, CYLINDRICAL_SHELL_PATH);

            // 3. Rename the Component
            string componentName = $"{CYLINDRICAL_SHELL_COMPONENT_NAME} {countNumber}";
            SWFeatureManager.GetFeatureByName(assemblyOfCylindricalShellsDoc, cylindricalShell.Name2).Name = componentName;

            // 4. Get Features for Mating
            Feature leftEndPlane = SWFeatureManager.GetFeatureByName(cylindricalShell, "Left End Plane");
            Feature rightEndPlane = SWFeatureManager.GetFeatureByName(cylindricalShell, "Right End Plane");
            Feature cylindricalShellCenterAxis = SWFeatureManager.GetFeatureByName(cylindricalShell, "Center Axis");

            
            // 5. Flip every second cylindrical shell
            bool flipDimension = false;
            if (countNumber % 2 == 0) flipDimension = true;

            Feature leftEndMate = null;
            Feature centerAxisMate = null;
            Feature frontPlaneMate = null;
            // 6. Create Mates
            try
            {
                leftEndMate = MateManager.CreateMate(
                componentFeature1: referenceCylindricalShell.GetRightEndPlane(),
                componentFeature2: leftEndPlane,
                alignmentType: MateAlignment.Aligned,
                name: $"{cylindricalShell.Name2} - {LEFT_END_PLANE_NAME}");

                centerAxisMate = MateManager.CreateMate(
                    componentFeature1: assemblyOfCylindricalShellsCenterAxis,
                    componentFeature2: cylindricalShellCenterAxis,
                    alignmentType: MateAlignment.Anti_Aligned,
                    name: $"{cylindricalShell.Name2} - {CENTER_AXIS_NAME}");


                //Mates new cylindrical shell's front plane with assembly's front plane with angle.
                frontPlaneMate = MateManager.CreateMate(
                    externalEntity: (Entity)assemblyOfCylindricalShellsFrontPlane,
                    componentEntity: (Entity)SWFeatureManager.GetMajorPlane(cylindricalShell, MajorPlane.Front),
                    referenceEntity: (Entity)assemblyOfCylindricalShellsCenterAxis,
                    angle: 0.78539816339744830961566084581988,
                    flipDimension: flipDimension,
                    name: $"{cylindricalShell.Name2} - {FRONT_PLANE_NAME}");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message );
            }

            // 7. Get cylindrical shell Entities and Initialize Settings
            _cylindricalShellSettings = new CylindricalShellSettings();
            try
            {
                GetCylindricalShellPIDs();

                // 8. Change diameter
                ChangeDiameter(diameter);

                // 9. Change length
                ChangeLength(cylindricalShell, length);
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting cylindrical shell entities: {ex.Message}");
            }

            void GetCylindricalShellPIDs()
            {
                try
                {
                    // Populate the _cylindricalShellSettings with the retrieved PIDs
                    _cylindricalShellSettings.PIDLeftEndPlane = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(leftEndPlane);
                    _cylindricalShellSettings.PIDRightEndPlane = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(rightEndPlane);
                    _cylindricalShellSettings.PIDCenterAxis = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(cylindricalShellCenterAxis);
                    _cylindricalShellSettings.PIDComponent = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(cylindricalShell);
                    _cylindricalShellSettings.PIDCenterAxisMate = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(centerAxisMate);
                    _cylindricalShellSettings.PIDLeftEndMate = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(leftEndMate);
                    _cylindricalShellSettings.PIDFrontPlaneMate = assemblyOfCylindricalShellsDoc.Extension.GetPersistReference3(frontPlaneMate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "The attribute could not be created.");
                    return;
                }
            }
        }

        /// <summary>
        /// DOCUMENT MUST BE ALREADY OPEN!!!
        /// Changes the diameter of a revolved feature (assuming a single relevant sub-feature).
        /// </summary>
        /// <param name="diameter"></param>
        private void ChangeDiameter(double diameter)
        {
            Component2 component = GetComponent();
            if (component == null)
            {
                throw new InvalidOperationException("Component not found.");
            }

            Feature revolveFeature = SWFeatureManager.GetFeatureByName(component, "Revolve");
            if (revolveFeature == null)
            {
                throw new InvalidOperationException("Revolve feature not found.");
            }

            Feature revolveSubFeature = revolveFeature.GetFirstSubFeature();
            if (revolveSubFeature == null)
            {
                throw new InvalidOperationException("Revolve sub-feature not found.");
            }

            revolveSubFeature.Parameter("Diameter").Value = diameter;
        }

        public Feature GetLeftEndPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDLeftEndPlane,
                        out int error);

        public Feature GetRightEndPlane() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDRightEndPlane,
                        out int error);

        public Component2 GetComponent() => (Component2)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDComponent,
                        out int error);

        public Feature GetCenterAxis() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDCenterAxis,
                        out int error);

        public Feature GetCenterAxisMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDCenterAxisMate,
                        out int error);

        public Feature GetLeftEndMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDLeftEndMate,
                        out int error);

        public Feature GetFrontPlaneMate() => (Feature)SolidWorksDocumentProvider.GetActiveDoc().Extension.GetObjectByPersistReference3(
                        _cylindricalShellSettings.PIDFrontPlaneMate,
                        out int error);

        /// <summary>
        /// Modifies the length of a cylindrical shell within a larger SolidWorks assembly.
        /// It achieves this by adjusting the distance of the reference plane associated with the right end of the shell.
        /// </summary>
        /// <param name="length"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void ChangeLength(Component2 cylindricalShellComp, double length)
        {
            // Get the model document
            ModelDoc2 assemblyOfCylindricalShellDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetCylindricalShellsAssemblyComponent().GetModelDoc2();
            if (assemblyOfCylindricalShellDoc == null)
            {
                throw new InvalidOperationException("Cylindrical shell assembly document not found.");
            }

            // Activate the assembly of cylindrical shells document
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfCylindricalShellDoc.GetTitle() + ".sldasm", true, 0, 0);

            // Change plane distance and rebuild the assembly
            try
            {
                // Get right end plane
                Feature rightEndPlane = GetRightEndPlane();

                // Activate cylindrical shell's document
                ModelDoc2 cylindricalShellModelDoc = cylindricalShellComp.GetModelDoc2();
                SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(cylindricalShellModelDoc.GetTitle(), true, 0, 0);

                SWFeatureManager.ChangeDistanceOfReferencePlane(SWFeatureManager.GetFeatureByName(SolidWorksDocumentProvider.GetActiveDoc(), rightEndPlane.Name), length);

                cylindricalShellModelDoc.EditRebuild3();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            assemblyOfCylindricalShellDoc.EditRebuild3();

            DocumentManager.UpdateAndSaveDocuments();
        }

        /// <summary>
        /// Changes angle of front plane mate
        /// </summary>
        /// <param name="angleInDegrees"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void ChangeAngle(double angleInDegrees)
        {
            // Get the model document (with null check)
            ModelDoc2 assemblyOfCylindricalShellDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetCylindricalShellsAssemblyComponent().GetModelDoc2();
            if (assemblyOfCylindricalShellDoc == null)
            {
                throw new InvalidOperationException("Assembly of cylindrical shells document not found.");
            }

            // Activate the shell assembly document
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfCylindricalShellDoc.GetTitle() + ".sldasm", true, 0, 0);

            MateFeatureData mateFeatureData = GetFrontPlaneMate().GetDefinition();
            AngleMateFeatureData angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;

            // Set the angle
            angleMateFeatureData.Angle = angleInDegrees * (Math.PI / 180);

            GetFrontPlaneMate().ModifyDefinition(angleMateFeatureData, assemblyOfCylindricalShellDoc, null);

            // Save and close assembly of cylindrical shells doc
            assemblyOfCylindricalShellDoc.Save3(
               (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
               (int)swFileSaveError_e.swGenericSaveError,
               (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfCylindricalShellDoc.GetTitle());

            // Save tank site assembly
            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        /// <summary>
        /// Deletes a cylindrical shell component from a SolidWorks assembly, including the associated file.
        /// </summary>
        public void Delete()
        {
            ModelDoc2 assemblyOfCylindricalShells = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr selectionManager = (SelectionMgr)assemblyOfCylindricalShells.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            GetComponent().Select4(false, selectData, false);

            //Get dished end document's path to delete the file
            ModelDoc2 componentDocument = GetComponent().GetModelDoc2();
            string path = componentDocument.GetPathName();

            //Delete selected dished end
            ((AssemblyDoc)assemblyOfCylindricalShells).DeleteSelections(0);

            //Rebuild assembly to release the file to be deleted
            assemblyOfCylindricalShells.EditRebuild3();

            //Delete the file
            File.Delete(path);
        }
    }
}

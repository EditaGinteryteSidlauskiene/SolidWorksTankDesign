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
    public class CylindricalShell
    {
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
            string projectFolder,
            CylindricalShell referenceCylindricalShell,
            Feature assemblyOfCylindricalShellsCenterAxis,
            Feature assemblyOfCylindricalShellsFrontPlane,
            double length,
            int countNumber)
        {
            // 1. Input Validation
            if (assemblyOfCylindricalShellsCenterAxis == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsCenterAxis));

            if (assemblyOfCylindricalShellsFrontPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsFrontPlane));

            if (referenceCylindricalShell == null)
                throw new ArgumentNullException(nameof(referenceCylindricalShell));

            ModelDoc2 assemblyOfCylindricalShellsDoc = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();
            if (assemblyOfCylindricalShellsDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");

            AddCylindricalShell(
                projectFolder,
                assemblyOfCylindricalShellsDoc,
                referenceCylindricalShell.GetRightEndPlane(),
                assemblyOfCylindricalShellsCenterAxis,
                assemblyOfCylindricalShellsFrontPlane,
                length,
                countNumber);
        }
        
        /// <summary>
        /// Adds cylindrical shell
        /// </summary>
        /// <param name="assemblyOfCylindricalShellsDoc"></param>
        /// <param name="referencePlane"></param>
        /// <param name="assemblyOfCylindricalShellsCenterAxis"></param>
        /// <param name="assemblyOfCylindricalShellsFrontPlane"></param>
        /// <param name="length"></param>
        /// <param name="countNumber"></param>
        private void AddCylindricalShell(
            string projectFolder,
            ModelDoc2 assemblyOfCylindricalShellsDoc,
            Feature referencePlane,
            Feature assemblyOfCylindricalShellsCenterAxis,
            Feature assemblyOfCylindricalShellsFrontPlane,
            double length,
            int countNumber)
        {
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;
            AssemblyOfCylindricalShells assemblyOfCylindricalShells = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells;

            string cylindricalShellPath = assemblyOfCylindricalShells._cylindricalShellDocPath;

            // Create a path where the cylindrical shell doc will be saved
            string targetPath = Path.Combine(projectFolder,
                $"Cylindrical shell " +
                $"{assemblyOfCylindricalShells.CylindricalShells.Count}.SLDPRT");

            // Open cylindrical shell doc and save it to a new destination
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(cylindricalShellPath);
            ModelDoc2 primaryCylindricalShellDoc = solidWorksApp.OpenDoc7(documentSpecification);

            primaryCylindricalShellDoc.SaveAs3(targetPath, 0, 0);

            // Close cylindrical shell doc and newly saved docs
            solidWorksApp.CloseDoc(primaryCylindricalShellDoc.GetTitle());
            solidWorksApp.CloseDoc(targetPath);

            // 2. Add and Make Independent Cylindrical Shell Component
            Component2 cylindricalShell = ComponentManager.AddComponentPart(targetPath);

            // 3. Get Features for Mating
            Feature leftEndPlane = SWFeatureManager.GetFeatureByName(cylindricalShell, "Left End Plane");
            Feature rightEndPlane = SWFeatureManager.GetFeatureByName(cylindricalShell, "Right End Plane");
            Feature cylindricalShellCenterAxis = SWFeatureManager.GetFeatureByName(cylindricalShell, "Center Axis");

            // 4. Flip every second cylindrical shell
            bool flipDimension = false;
            if (countNumber % 2 == 0) flipDimension = true;

            Feature leftEndMate = null;
            Feature centerAxisMate = null;
            Feature frontPlaneMate = null;
            // 5. Create Mates
            try
            {
                leftEndMate = MateManager.CreateMate(
                componentFeature1: referencePlane,
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            // 6. Get cylindrical shell Entities and Initialize Settings
            _cylindricalShellSettings = new CylindricalShellSettings();
            try
            {
                GetCylindricalShellPIDs();

                // 7. Change length
                ChangeLength(length);
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
                MessageBox.Show("Unable to change diameter.");
                return;
            }

            Feature revolveFeature = SWFeatureManager.GetFeatureByName(component, "Revolve");
            if (revolveFeature == null)
            {
                MessageBox.Show("Unable to change diameter.");
                return;
            }

            Feature revolveSubFeature = revolveFeature.GetFirstSubFeature();
            if (revolveSubFeature == null)
            {
                MessageBox.Show("Unable to change diameter.");
                return;
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

        public void CompleteFirstCylindricalShel(string projectFolder, double length)
        {
            AssemblyOfCylindricalShells assemblyOfCylindricalShells = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells;
            
            ModelDoc2 assemblyOfCylindricalShellsDoc = assemblyOfCylindricalShells.ActivateDocument();
            if (assemblyOfCylindricalShellsDoc == null)
                throw new InvalidOperationException("Active SolidWorks document not found.");
            

            Feature assemblyOfCylindricalShellsCenterAxis = assemblyOfCylindricalShells.GetCenterAxis();
            Feature assemblyOfCylindricalShellsFrontPlane = SWFeatureManager.GetMajorPlane(assemblyOfCylindricalShellsDoc, MajorPlane.Front);
            Feature assemblyOfCylindricalShellsRightPlane = SWFeatureManager.GetMajorPlane(assemblyOfCylindricalShellsDoc, MajorPlane.Right);
            // 1. Input Validation
            if (assemblyOfCylindricalShellsCenterAxis == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsCenterAxis));

            if (assemblyOfCylindricalShellsFrontPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsFrontPlane));

            if (assemblyOfCylindricalShellsRightPlane == null)
                throw new ArgumentNullException(nameof(assemblyOfCylindricalShellsRightPlane));

            AddCylindricalShell(
                projectFolder,
                assemblyOfCylindricalShellsDoc,
                assemblyOfCylindricalShellsRightPlane,
                assemblyOfCylindricalShellsCenterAxis,
                assemblyOfCylindricalShellsFrontPlane,
                length,
                1);
        }

        /// <summary>
        /// Modifies the length of a cylindrical shell within a larger SolidWorks assembly.
        /// It achieves this by adjusting the distance of the reference plane associated with the right end of the shell.
        /// </summary>
        /// <param name="length"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void ChangeLength(double length)
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
                //ModelDoc2 cylindricalShellModelDoc = cylindricalShellComp.GetModelDoc2();
                ModelDoc2 cylindricalShellModelDoc = GetComponent().GetModelDoc2();
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

            // SaveInitialConfiguration and close assembly of cylindrical shells doc
            assemblyOfCylindricalShellDoc.Save3(
               (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
               (int)swFileSaveError_e.swGenericSaveError,
               (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfCylindricalShellDoc.GetTitle());

            // SaveInitialConfiguration tank site assembly
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

        /// <summary>
        /// Retrieves the diameter of the cylindrical shell from the SolidWorks model.
        /// </summary>
        /// <returns>The diameter of the cylindrical shell in millimeters. Returns 0 if an error occurs.</returns>
        public double GetCylindricalShellDiameter()
        {
            // Activate the assembly cylindrical shells document.
            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();

            // Get the cylindrical shell component.
            Component2 cylindricalShellComp = GetComponent();
            if (cylindricalShellComp == null)
            {
                // Handle the case where the component is not found.
                return 0;
            }

            // Get the "Revolve" feature from the cylindrical shell component.
            Feature revolveFeature = SWFeatureManager.GetFeatureByName(cylindricalShellComp, "Revolve");
            if (revolveFeature == null)
            {
                // Handle the case where the "Revolve" feature is not found.
                return 0;
            }

            // Get the first sub-feature of the "Revolve" feature. This is assumed to contain the diameter parameter.
            Feature revolveSubFeature = revolveFeature.GetFirstSubFeature();
            if (revolveSubFeature == null)
            {
                // Handle the case where the sub-feature is not found.
                return 0;
            }

            // Get the value of the "Diameter" parameter from the sub-feature.
            double diameter = revolveSubFeature.Parameter("Diameter").Value;

            // Close the cylindrical shell assembly document.
            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CloseDocument();

            // Return the extracted diameter.
            return diameter;
        }
    }
}

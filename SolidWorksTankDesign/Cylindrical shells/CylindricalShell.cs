using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using System.Windows.Forms;
using System;
using System.IO;
using System.Xml.Linq;
using AddinWithTaskpane;
using SolidWorks.Interop.swconst;

namespace SolidWorksTankDesign
{
    internal class CylindricalShell
    {
        private const string CYLINDRICAL_SHELL_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\Shell Cylyndrical ø1600×6 L1000.SLDPRT";
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

        public CylindricalShell(
            CylindricalShell referenceCylindricalShell,
            Feature assemblyOfCylindricalShellsCenterAxis,
            Feature assemblyOfCylindricalShellsFrontPlane,
            double length,
            double diameter,
            int countNumber)
        {
            ModelDoc2 assemblyOfCylindricalShellsDoc = SolidWorksDocumentProvider.GetActiveDoc();
            
            Component2 cylindricalShell = ComponentManager.AddComponentPart(referenceCylindricalShell.GetComponent().GetPathName());
            
            ComponentManager.MakeCylindricalShellIndependent(cylindricalShell, CYLINDRICAL_SHELL_PATH);

            string componentName = $"{CYLINDRICAL_SHELL_COMPONENT_NAME} {countNumber}";
            Feature cylindricalShellFeature =  FeatureManager.GetFeatureByName(assemblyOfCylindricalShellsDoc, cylindricalShell.Name2);
            cylindricalShellFeature.Name = componentName;

            Feature leftEndPlane = FeatureManager.GetFeatureByName(cylindricalShell, "Left End Plane");
            Feature rightEndPlane = FeatureManager.GetFeatureByName(cylindricalShell, "Right End Plane");
            Feature cylindricalShellCenterAxis = FeatureManager.GetFeatureByName(cylindricalShell, "Center Axis");

            Feature leftEndMate = MateManager.CreateMate(
                componentFeature1: referenceCylindricalShell.GetRightEndPlane(),
                componentFeature2: leftEndPlane,
                alignmentType: MateAlignment.Aligned,
                name: $"{cylindricalShell.Name2} - {LEFT_END_PLANE_NAME}");

            Feature centerAxisMate = MateManager.CreateMate(
                componentFeature1: assemblyOfCylindricalShellsCenterAxis,
                componentFeature2: cylindricalShellCenterAxis,
                alignmentType: MateAlignment.Anti_Aligned,
                name: $"{cylindricalShell.Name2} - {CENTER_AXIS_NAME}");

            //Flip every second cylindrical shell
            bool flipDimension = false;
            if (countNumber % 2 == 0) flipDimension = true;

            //Mates new cylindrical shell's front plane with assembly's front plane with angle.
            Feature frontPlaneMate = MateManager.CreateMate(
                ExternalEntity: (Entity)assemblyOfCylindricalShellsFrontPlane,
                ComponentEntity: (Entity)FeatureManager.GetMajorPlane(cylindricalShell, MajorPlane.Front),
                ReferenceEntity: (Entity)assemblyOfCylindricalShellsCenterAxis,
                Angle: 0.78539816339744830961566084581988,
                FlipDimension: flipDimension,
                Name: $"{cylindricalShell.Name2} - {FRONT_PLANE_NAME}");

            
            //assemblyOfCylindricalShellsDoc.EditRebuild3();

            _cylindricalShellSettings = new CylindricalShellSettings();
            try
            {
                GetCylindricalShellPIDs();
            }
            catch (Exception ex)
            {
                // Display a user-friendly error message
                MessageBox.Show($"Error getting cylindrical shell entities: {ex.Message}");
            }

            ChangeDiameter(diameter);

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
        /// </summary>
        /// <param name="diameter"></param>
        private void ChangeDiameter(double diameter)
        {
            //Get Revolve feature
            Feature revolveFeature = FeatureManager.GetFeatureByName(GetComponent(), "Revolve");

            Feature childFeature = revolveFeature.GetFirstSubFeature();
            childFeature.Parameter("Diameter").Value = diameter;
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

        public void ChangeLength(double length)
        {
            ModelDoc2 assemblyOfCylindricalShellDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetCylindricalShellsAssemblyComponent().GetModelDoc2();
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfCylindricalShellDoc.GetTitle() + ".sldasm", true, 0, 0);

            FeatureManager.ChangeDistanceOfReferencePlane(GetRightEndPlane(), length);

            assemblyOfCylindricalShellDoc.EditRebuild3();

            assemblyOfCylindricalShellDoc.Save3(
               (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
               (int)swFileSaveError_e.swGenericSaveError,
               (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfCylindricalShellDoc.GetTitle());

            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
               (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
               (int)swFileSaveError_e.swGenericSaveError,
               (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        public void ChangeAngle(double angleInDegrees)
        {
            ModelDoc2 assemblyOfCylindricalShellDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetCylindricalShellsAssemblyComponent().GetModelDoc2();
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(assemblyOfCylindricalShellDoc.GetTitle() + ".sldasm", true, 0, 0);

            MateFeatureData mateFeatureData = GetFrontPlaneMate().GetDefinition();
            AngleMateFeatureData angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;

            angleMateFeatureData.Angle = angleInDegrees * (Math.PI / 180);

            GetFrontPlaneMate().ModifyDefinition(angleMateFeatureData, assemblyOfCylindricalShellDoc, null);

            assemblyOfCylindricalShellDoc.Save3(
               (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
               (int)swFileSaveError_e.swGenericSaveError,
               (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
            SolidWorksDocumentProvider._solidWorksApplication.CloseDoc(assemblyOfCylindricalShellDoc.GetTitle());

            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

       

        /// <summary>
        /// Deletes the cylindrical shell component
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

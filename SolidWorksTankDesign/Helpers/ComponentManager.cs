using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.Helpers;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal static class ComponentManager
    {
        /// <summary>
        /// Adds a component part to a SolidWorks assembly document
        /// </summary>
        /// <param name="solidWorksApplication"></param>
        /// <param name="assemblyDocument"></param>
        /// <param name="componentPath"></param>
        /// <returns></returns>
        public static Component2 AddComponentPart(string componentPath)
        {
            SldWorks solidWorksApplication = SolidWorksDocumentProvider._solidWorksApplication;
            ModelDoc2 assemblyDocument = SolidWorksDocumentProvider.GetActiveDoc();
            // Open the document of the component to be added
            solidWorksApplication.OpenDoc6(componentPath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 1, 1);

            // Add the part to the assembly document
            Component2 Component = ((AssemblyDoc)assemblyDocument).AddComponent5(componentPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 1, 0, 0);

            // Close the document of added component
            solidWorksApplication.CloseDoc(componentPath);

            // Making the added component float
            // Checking if the component is fixed
            if (Component.IsFixed() == true)
            {
                // Selecting the component
                Component.Select2(false, 0);

                // Unfixing component
                ((AssemblyDoc)assemblyDocument).UnfixComponent();
            }

            return Component;
        }

        /// <summary>
        /// Adds Component assembly. First, the parent document has to be activated.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="assemblyDocument"></param>
        /// <param name="componentPath"></param>
        /// <returns></returns>
        public static Component2 AddComponentAssembly(ModelDoc2 assemblyDocument, string componentPath)
        {
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;

            int error = 0;
            int warning = 0;
            // Open the document of the component to be added
            DocumentSpecification documentSpecification = solidWorksApp.GetOpenDocSpec(componentPath);
            ModelDoc2 componentDoc = solidWorksApp.OpenDoc7(documentSpecification);

            //ModelDoc2 componentDoc = solidWorksApp.OpenDoc6(componentPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref error, ref warning);
            
            if(componentDoc is null)
            {
                string nozzleDocTitle = componentPath.Split('\\').Last().Split('.')[0];
                object[] activeDocs = solidWorksApp.GetDocuments();
                foreach (object activeDoc in activeDocs)
                {
                    string nameDoc = ((ModelDoc2)activeDoc).GetTitle();
                    string activeDocPath = ((ModelDoc2)activeDoc).GetPathName();
                    if (nameDoc == nozzleDocTitle &&
                        activeDocPath == componentPath)
                    {
                        componentDoc = (ModelDoc2)activeDoc;
                        solidWorksApp.ActivateDoc3(componentDoc.GetTitle(), true, 0, 0);
                        break;
                    }
                }
            }
            // Add the part to the assembly document
            Component2 component = ((AssemblyDoc)assemblyDocument).AddComponent5(componentPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);

            if (component == null)
            {
                return null;
            }

            // Close the document of added component
            solidWorksApp.CloseDoc(componentPath);

            // Making the added component float
            // Checking if the component is fixed
            if (component.IsFixed() == true)
            {
                // Selecting the component
                component.Select2(false, 0);

                // Unfixing component
                ((AssemblyDoc)assemblyDocument).UnfixComponent();
            }

            return component;
        }

        /// <summary>
        /// Makes main (left or right) dished end independent. Assembly of dished ends must be already active
        /// </summary>
        /// <param name="componentPart"></param>
        /// <param name="newDocumentPath"></param>
        /// <returns></returns>
        public static bool MakeMainDishedEndIndependent(Component2 componentPart, string newDocumentPath)
        {
            // Get assembly of dished ends doc
            ModelDoc2 assemblyDocument = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr swSelMgr = (SelectionMgr)assemblyDocument.SelectionManager;
            SelectData swSelData = swSelMgr.CreateSelectData();

            //Select component 
            componentPart.Select4(false, swSelData, false);

            //Make the component part independent
            bool status = ((AssemblyDoc)assemblyDocument).MakeIndependent(newDocumentPath);

            assemblyDocument.ClearSelection2(true);

            return status;
        }

        /// <summary>
        /// Makes a new dished end component part independent by creating a new file and renames the component in the assembly.
        /// </summary>
        /// <param name="componentPart"></param>
        /// <param name="componentPartPath"></param>
        /// <returns></returns>
        public static bool MakeDishedEndIndependent(Component2 componentPart, string componentPartPath)
        {
            ModelDoc2 assemblyDocument = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr swSelMgr = (SelectionMgr)assemblyDocument.SelectionManager;
            SelectData swSelData = swSelMgr.CreateSelectData();

            //Select component part
            componentPart.Select4(false, swSelData, false);

            //Get the number of component in the assembly.
            int componentNumber = ((IAssemblyDoc)assemblyDocument).GetComponents(true).Length-2;

            //Add component path and number of ticks of current time into component part path. This path will be used to create a new document of the component.
            componentPartPath = componentPartPath.Insert(
                componentPartPath.IndexOf('.'),
                "_" + componentNumber.ToString() + "_" + DateTime.Now.Ticks);

            //Make the component part independent
            bool status = ((AssemblyDoc)assemblyDocument).MakeIndependent(componentPartPath);

            assemblyDocument.ClearSelection2(true);

            return true;
        }

        /// <summary>
        /// Makes a new cylindrical shell component part independent by creating a new file and renames the component in the assembly.
        /// </summary>
        /// <param name="componentPart"></param>
        /// <param name="componentPartPath"></param>
        /// <returns></returns>
        public static bool MakeCompartmentIndependent(Component2 componentPart, string compartmentDocPath)
        {
            ModelDoc2 assemblyDocument = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr swSelMgr = (SelectionMgr)assemblyDocument.SelectionManager;
            SelectData swSelData = swSelMgr.CreateSelectData();

            //Select component part
            componentPart.Select4(false, swSelData, false);

            //Get the number of component in the assembly.
            int componentNumber = ((IAssemblyDoc)assemblyDocument).GetComponents(true).Length;

            //Make the component part independent
            bool status = ((AssemblyDoc)assemblyDocument).MakeIndependent(compartmentDocPath);

            assemblyDocument.ClearSelection2(true);

            return true;
        }

        /// <summary>
        /// Makes a new cylindrical shell component part independent by creating a new file and renames the component in the assembly.
        /// </summary>
        /// <param name="componentPart"></param>
        /// <param name="componentPartPath"></param>
        /// <returns></returns>
        public static bool MakeComponentIndependent(Component2 componentPart, string componentPartPath)
        {
            ModelDoc2 assemblyDocument = SolidWorksDocumentProvider.GetActiveDoc();

            SelectionMgr swSelMgr = (SelectionMgr)assemblyDocument.SelectionManager;
            SelectData swSelData = swSelMgr.CreateSelectData();

            //Select component part
            componentPart.Select4(false, swSelData, false);

            //Get the number of component in the assembly.
            int componentNumber = ((IAssemblyDoc)assemblyDocument).GetComponents(true).Length;
            
            //Add component path and number of ticks of current time into component part path. This path will be used to create a new document of the component.
            componentPartPath = componentPartPath.Insert(
                componentPartPath.IndexOf('.'),
                "_" + componentNumber.ToString() + "_" + DateTime.Now.Ticks);

            //Make the component part independent
            bool status = ((AssemblyDoc)assemblyDocument).MakeIndependent(componentPartPath);

            assemblyDocument.ClearSelection2(true);

            return true;
        }

        /// <summary>
        /// Gets or sets the alignment of the dished end. If the alignment changes, a method to modify the existing alignment is triggered.
        /// </summary>
        public static void SetAlignment(
            Component2 dishedEndComponent, 
            DishedEndAlignment dishedEndAlignment,
            Feature centerAxisMate,
            Feature rightPlaneMate,
            DishedEndSettings dishedEndSettings)
        {
            if (dishedEndAlignment != GetAlignment(dishedEndComponent))
            {
                ChangeAlignment(centerAxisMate, rightPlaneMate, dishedEndSettings);
            }
        }

        /// <summary>
        /// Gets whether dished end is alligned left or right.
        /// </summary>
        /// <returns></returns>
        private static DishedEndAlignment GetAlignment(Component2 dishedEndComponent)
        {
            //Get the transformation matrix from the dishedEnd object
            MathTransform transform = dishedEndComponent.Transform2;

            //Transform the reference point (1, 0, 0) using the transformation matrix
            double[] TransformedVector = MathUtility.TransformVector(SolidWorksDocumentProvider._solidWorksApplication, transform, new double[3] { 1, 0, 0 });

            //Determine orientation based on the transformed point's X-coordinate
            return (TransformedVector[0] > 0 ? DishedEndAlignment.Left : DishedEndAlignment.Right);
        }

        private static void UpdateAxisAndRightPlaneMatesPIDs(Feature centerAxisMate, Feature rightPlaneMate, DishedEndSettings dishedEndSettings)
        {
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider.GetActiveDoc();
            dishedEndSettings.PIDCenterAxisMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(centerAxisMate);
            dishedEndSettings.PIDRightPlaneMate = assemblyOfDishedEndsDoc.Extension.GetPersistReference3(rightPlaneMate);
        }

        /// <summary>
        /// Changes alignment of the dished end component
        /// </summary>
        public static void ChangeAlignment(Feature centerAxisMate, Feature rightPlaneMate, DishedEndSettings dishedEndSettings)
        {
            try
            {
                SWFeatureManager.Suppress(centerAxisMate);

                //Change alignment of the component
                //Warning message if ChangeAlignement() did not work
                if (!MateManager.ChangeAlignment(rightPlaneMate))
                {
                    MessageBox.Show("Failed to change the right plane mate alignment.");
                    SWFeatureManager.Unsuppress(centerAxisMate);
                    return;
                }

                //Change alignment of axis
                //Warning message if ChangeAlignement() did not work
                if (!MateManager.ChangeAlignment(centerAxisMate))
                {
                    MessageBox.Show("Failed to change the center axis mate alignment.");
                    MateManager.ChangeAlignment(rightPlaneMate);
                    SWFeatureManager.Unsuppress(centerAxisMate);
                    return;
                }

                //Unsuppress axis mate
                SWFeatureManager.Unsuppress(centerAxisMate);

                UpdateAxisAndRightPlaneMatesPIDs(centerAxisMate, rightPlaneMate, dishedEndSettings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Mate alignment could not be changed." + ex.Message);
                return;
            }
        }

        public static void RefreshDishedEnds()
        {
            ModelDoc2 activeDoc = SolidWorksDocumentProvider.GetActiveDoc();

            foreach (DishedEnd dishedEnd in SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds)
            {
                Feature positionPlane = dishedEnd.GetPositionPlane();

                RefPlaneFeatureData positionPlaneFeatureData = positionPlane.GetDefinition();
                positionPlane.ModifyDefinition(positionPlaneFeatureData, activeDoc, null);
            }

            Feature rightDishedEndPositionPlane = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.GetPositionPlane();

            RefPlaneFeatureData rightDishedEndPositionPlaneFeatureData = rightDishedEndPositionPlane.GetDefinition();
            rightDishedEndPositionPlane.ModifyDefinition(rightDishedEndPositionPlaneFeatureData, activeDoc, null);
        }
    }
}

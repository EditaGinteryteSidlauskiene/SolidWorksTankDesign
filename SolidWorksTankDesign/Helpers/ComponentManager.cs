using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;

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
        public static bool MakeCylindricalShellIndependent(Component2 componentPart, string componentPartPath)
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


    }
}

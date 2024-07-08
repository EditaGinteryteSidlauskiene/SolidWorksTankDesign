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
        public static Component2 AddComponentPart(SldWorks solidWorksApplication, ModelDoc2 assemblyDocument, string componentPath)
        {
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
        /// Makes a new component part independent by creating a new file and renames the component in the assembly.
        /// </summary>
        /// <param name="AssemblyDocument"></param>
        /// <param name="ComponentPart"></param>
        /// <param name="ComponentPartPath"></param>
        /// <param name="ComponentPartName"></param>
        /// <returns></returns>
        public static bool MakeComponentPartIndependent(ModelDoc2 assemblyDocument, Component2 componentPart, string componentPartPath)
        {

            SelectionMgr swSelMgr = (SelectionMgr)assemblyDocument.SelectionManager;
            SelectData swSelData = swSelMgr.CreateSelectData();

            //Select component part
            componentPart.Select4(false, swSelData, false);

            //Get the number of component in the assembly.
            int componentNumber = ((IAssemblyDoc)assemblyDocument).GetComponents(true).Length-2;

            //Add component number and ticks of current time into component part path. This path will be used to create a new document of the component.
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

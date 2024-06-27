using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using WarningAndErrorService;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    /* This is the highest class in the hierarchy. At least, it must have the center axis and 
     Tank component.*/
    internal class TankSiteAssembly
    {
        SldWorks solidWorksApplication;
        private readonly ModelDoc2 tankSiteModelDoc;
        private readonly AssemblyDoc tankSiteAssemblyDoc;
        private readonly WarningService warningService;

        public FeatureAxis tankSiteCenterAxis;

        /// <summary>
        /// TankSiteAssembly constructor. 
        /// </summary>
        /// <param name="WarningService"></param>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public TankSiteAssembly(WarningService WarningService, SldWorks SolidWorksApplication, ModelDoc2 TankSiteModelDoc)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in TankSideAssembly constructor.");
                return;
            }

            //If the document is null, the method terminates.
            if (TankSiteModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in TankSideAssembly constructor.");
                return;
            }

            solidWorksApplication = SolidWorksApplication;
            tankSiteModelDoc = TankSiteModelDoc;
            tankSiteAssemblyDoc = (AssemblyDoc)tankSiteModelDoc;

            //// Attempt to initialize the tank site assembly
            //if (!TrySetTankSite())
            //{
            //    warningService.AddError("The TankSideAssembly could not be set properly.");
            //    return;
            //}
        }

        private bool TrySetTankSite()
        {
            // Check if there is the center axis, and if it correctly defined
            //Check if there is one center axis
            if (!UtilitiesCheck.IsNumberOfReferenceAxisCorrect(warningService, tankSiteModelDoc, 1, out List<FeatureAxis> CenterAxisList)) return false;

            //Get all assemblies, check them and assign
            List<Component2> listOfAssemblies = Utilities.GetTopLevelComponents(tankSiteModelDoc);

            //Check if there are at least 1 components in the Tank Site Assembly
            if (listOfAssemblies.Count < 1)
            {
                warningService.AddError("Incorrect number of components in the Tank Site Assembly. Min number of components = 1");
                return false;
            }

            ////Get ModelDoc2 of each assembly and assign Shell 
            //foreach (Component2 assembly in listOfAssemblies)
            //{
            //    //Activate current assemlby's document
            //    tankDocument = Utilities.ActivateDocumentOfComponent(solidWorksApplication, assembly);

            //    //Try to assign tank assembly and close the document

            //    if (TrySetTank(tankDocument))
            //    {
            //        tankAsComponent = assembly;
            //        continue;
            //    }

            //    else
            //    {
            //        solidWorksApplication.CloseDoc(tankDocument.GetTitle());
            //        continue;
            //    }

            //}

            //if (tank == null)
            //{
            //    warningService.AddError("Incorrect Tank Assembly. Assembly of Shell was not assigned.");
            //    return false;
            //}

            // Assign Tank Site Assembly's center axis
            tankSiteCenterAxis = CenterAxisList[0];

            return true;
        }
    }
}

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
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
        private readonly ModelDocExtension tankSiteModelDocExt;
        private readonly AssemblyDoc tankSiteAssemblyDoc;
        private readonly WarningService warningService;

        private List<(string ParameterName, string ParameterValue)> parametersList = new List<(string ParameterName, string ParameterValue)> ();
        public string persistentReferenceIdAttributeName = "PersistentReferenceIDs";

        public FeatureAxis tankSiteCenterAxis;


        /// <summary>
        /// Initializes a TankSiteAssembly object, representing a SolidWorks tank site assembly model.It ensures 
        /// the provided SolidWorks application and model document are valid and performs initial setup operations. 
        /// </summary>
        /// <param name="warningService"></param>
        /// <param name="solidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public TankSiteAssembly(WarningService warningService, SldWorks solidWorksApplication, ModelDoc2 tankSiteModelDoc)
        {
            this.warningService = warningService;

            // Null checks
            if (solidWorksApplication == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly.solidWorksApplication), "SolidWorks application is required.");
            }

            if (tankSiteModelDoc == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly.tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the SolidWorks application and model objects
            this.solidWorksApplication = solidWorksApplication;
            this.tankSiteModelDoc = tankSiteModelDoc;

            // Get the model document extension for additional functionality
            this.tankSiteModelDocExt = this.tankSiteModelDoc.Extension;

            // Cast the model document to an AssemblyDoc
            this.tankSiteAssemblyDoc = (AssemblyDoc)this.tankSiteModelDoc;

            //---------------- Initialization -------------------------------------

            // Set Tank Site Center Axis Presistent Reference ID
            AddTankSiteAssemblyCenterAxisID();

            // Set Tank Workshop Assembly (as Component) Presistent Reference ID
            AddTankWorkshopAssemblyComponentID();

            AttributeManager.CreateAttribute(this.solidWorksApplication, this.tankSiteModelDoc, this.tankSiteModelDoc, parametersList, persistentReferenceIdAttributeName);
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the center axis feature within the tank site assembly model. 
        /// </summary>
        private void AddTankSiteAssemblyCenterAxisID()
        {
            // Get center axis
            Feature tankSiteAssemblyCenterAxis = Utilities.GetNthFeatureOfType(tankSiteModelDoc, FeatureType.RefAxis, 1);

            // Safety check: Ensure the center axis feature was found
            if (tankSiteAssemblyCenterAxis == null)
            {
                throw new InvalidOperationException("Center axis feature not found in the tank site assembly.");
            }

            // Get the persistent reference ID of the center axis
            byte[] tankSiteAssemblyCenterAxisID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, tankSiteAssemblyCenterAxis);

            string persistentReferenceIdString = string.Join("-", tankSiteAssemblyCenterAxisID);

            parametersList.Add(("P1", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the tank workshop assembly component within the tank site assembly model.
        /// </summary>
        private void AddTankWorkshopAssemblyComponentID()
        {
            // Get TankWorkshopAssembly as Component
            Component2 tankWorkshopAssembly = Utilities.GetAllComponents(tankSiteModelDoc)[0].Component;

            // Safety check: Ensure the tank workshop assembly component was found
            if (tankWorkshopAssembly == null)
            {
                throw new InvalidOperationException("Tank workshop assembly component not found in the tank site assembly.");
            }

            // Get the persistent reference ID of the tank workshop assembly component
            byte[] tankWorkshopAssemblyID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, tankWorkshopAssembly);

            string persistentReferenceIdString = string.Join("-", tankWorkshopAssemblyID);

            parametersList.Add(("P2", persistentReferenceIdString));
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

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
        public string persistentReferenceIdAttributeName = "TankSiteAssemblyPersistentReferenceIDs";

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

            try
            {
                AddTankSiteAssemblyPersistantReferenceIds();
            }
            catch (Exception ex) { }

            AttributeManager.CreateAttribute(this.solidWorksApplication, this.tankSiteModelDoc, this.tankSiteModelDoc, parametersList, persistentReferenceIdAttributeName);
        }

        /// <summary>
        /// Adds persistent reference IDs to key components and features within a tank site assembly model.
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        private void AddTankSiteAssemblyPersistantReferenceIds()
        {
            // Get components
            Component2 tankWorkshopAssemblyAsComponent = Utilities.GetAllComponents(tankSiteModelDoc)[0].Component;
            Component2 tankAssemblyAsComponent = Utilities.GetAllComponents(tankWorkshopAssemblyAsComponent)[0].Component;
            Component2 shellAssemblyAsComponent = Utilities.GetAllComponents(tankAssemblyAsComponent)[0].Component;

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

            // Add persistent reference IDs
            AddTankSiteAssemblyCenterAxisID();
            AddTankWorkshopAssemblyComponentID(tankWorkshopAssemblyAsComponent);
            AddTankWorkshopAssemblyCenterAxisMateID(tankWorkshopAssemblyAsComponent);
            AddTankAssemblyComponentID(tankAssemblyAsComponent);
            AddShellAssemblyComponentID(shellAssemblyAsComponent);
            AddAssemblyOfDishedEndsComponentID(shellAssemblyAsComponent);
            AddAssemblyOfCylindricalShellsComponentID(shellAssemblyAsComponent);
            AddCompartmentsAssemblyComponentID(shellAssemblyAsComponent);
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

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", tankSiteAssemblyCenterAxisID);

            parametersList.Add(("P1", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the tank workshop assembly component within the tank site assembly model.
        /// </summary>
        private void AddTankWorkshopAssemblyComponentID(Component2 tankWorkshopAssemblyAsComponent)
        {
            // Get the persistent reference ID of the tank workshop assembly component
            byte[] tankWorkshopAssemblyID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, tankWorkshopAssemblyAsComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", tankWorkshopAssemblyID);

            parametersList.Add(("P2", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the tank workshop assembly's center axis mate within the tank site assembly model.
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        private void AddTankWorkshopAssemblyCenterAxisMateID(Component2 tankWorkshopAssemblyAsComponent)
        {
            // Get the mate
            Mate2 tankWorkshopAssemblyCenterAxisMate = tankWorkshopAssemblyAsComponent.GetMates()[0];

            // Safety check: Ensure the center axis feature was found
            if (tankWorkshopAssemblyCenterAxisMate == null)
            {
                throw new InvalidOperationException("Center axis mate not found in the tank site assembly.");
            }

            // Get the persistent reference ID of the tank workshop assembly component
            byte[] tankWorkshopAssemblyCenterAxisMateID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, tankWorkshopAssemblyCenterAxisMate);

            string persistentReferenceIdString = string.Join("-", tankWorkshopAssemblyCenterAxisMateID);

            parametersList.Add(("P3", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the tank assembly component within within the tank site assembly model.
        /// </summary>
        private void AddTankAssemblyComponentID(Component2 tankAssemblyAsComponent)
        {
            // Safety check: Ensure the tank assembly component was found
            if (tankAssemblyAsComponent == null)
            {
                throw new InvalidOperationException("Tank assembly component not found in the tank workshop assembly.");
            }

            // Get the persistent reference ID of the tank assembly component
            byte[] tankAssemblyID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, tankAssemblyAsComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", tankAssemblyID);

            parametersList.Add(("P4", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the shell assembly component within the tank site assembly model.
        /// </summary>
        private void AddShellAssemblyComponentID(Component2 shellAssemblyAsComponent)
        {
            // Get the persistent reference ID of the shell assembly component
            byte[] shellAssemblyID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, shellAssemblyAsComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", shellAssemblyID);

            parametersList.Add(("P5", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the dished ends assembly component within the tank site assembly model.
        /// </summary>
        /// <param name="shellAssemblyAsComponent"></param>
        private void AddAssemblyOfDishedEndsComponentID(Component2 shellAssemblyAsComponent)
        {
            Component2 assemblyOfDishedEndsComponent = Utilities.GetAllComponents(shellAssemblyAsComponent)[0].Component;

            // Get the persistent reference ID of the shell assembly component
            byte[] assemblyOfDishedEndsComponentID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, assemblyOfDishedEndsComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", assemblyOfDishedEndsComponentID);

            parametersList.Add(("P6", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the cylindrical shells assembly component within the tank site assembly model.
        /// </summary>
        /// <param name="shellAssemblyAsComponent"></param>
        private void AddAssemblyOfCylindricalShellsComponentID(Component2 shellAssemblyAsComponent)
        {
            Component2 assemblyOfCylindricalShellsComponent = Utilities.GetAllComponents(shellAssemblyAsComponent)[1].Component;

            // Get the persistent reference ID of the shell assembly component
            byte[] assemblyOfCylindricalShellsComponentID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, assemblyOfCylindricalShellsComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", assemblyOfCylindricalShellsComponentID);

            parametersList.Add(("P7", persistentReferenceIdString));
        }

        /// <summary>
        /// This method retrieves and stores the persistent reference ID of the compartments assembly component within the tank site assembly model.
        /// </summary>
        /// <param name="shellAssemblyAsComponent"></param>
        private void AddCompartmentsAssemblyComponentID(Component2 shellAssemblyAsComponent)
        {
            Component2 compartmentsAssemblyComponent = Utilities.GetAllComponents(shellAssemblyAsComponent)[2].Component;

            // Get the persistent reference ID of the shell assembly component
            byte[] compartmentsAssemblyComponentID = Utilities.GetPersistentReferenceId(tankSiteModelDocExt, compartmentsAssemblyComponent);

            // Convert persistent reference id into string
            string persistentReferenceIdString = string.Join("-", compartmentsAssemblyComponentID);

            parametersList.Add(("P8", persistentReferenceIdString));
        }
    }
}

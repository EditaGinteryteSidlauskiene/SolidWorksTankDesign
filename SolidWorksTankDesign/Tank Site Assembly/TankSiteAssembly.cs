﻿using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    /* This is the highest class in the hierarchy. At least, it must have the center axis and 
     Tank component.*/
    internal class TankSiteAssembly
    {
        [JsonIgnore]
        public ModelDoc2 _tankSiteModelDoc;
        [JsonIgnore]
        private ModelDocExtension _tankSiteDocExtension => _tankSiteModelDoc.Extension;
        
        public TankSiteAssemblySettings _tankSiteAssemblySettings;
        public CompartmentsManager _compartmentsManager;
        public AssemblyOfDishedEnds _assemblyOfDishedEnds;
        public AssemblyOfCylindricalShells _assemblyOfCylindricalShells;

        /// <summary>
        /// DO NOT DELETE IT!!! This constructor is needed for Json deserializer.
        /// </summary>
        public TankSiteAssembly()
        {
            // Null checks
            if (SolidWorksDocumentProvider._solidWorksApplication == null)
            {
                throw new ArgumentNullException(nameof(SolidWorksDocumentProvider._solidWorksApplication), "SolidWorks application is required.");
            }
            if (SolidWorksDocumentProvider.GetActiveDoc() == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the model doc object
            _tankSiteModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Create a default instance of the TankSiteAssemblySettings class
            // to hold settings for the TankSiteAssembly object.
            _tankSiteAssemblySettings = new TankSiteAssemblySettings();
            _compartmentsManager = new CompartmentsManager();
            _assemblyOfDishedEnds = new AssemblyOfDishedEnds();
            _assemblyOfCylindricalShells = new AssemblyOfCylindricalShells();
        }

        /// <summary>
        /// Initializes the tank site assembly's settings and creates a configuration attribute containing those settings.
        /// This method is called when new attribute needs to be created.
        /// </summary>
        /// <exception cref="ArgumentNullException"></exception>
        public void InitializeAndStoreTankSiteConfiguration()
        {
            // Null checks
            if (SolidWorksDocumentProvider._solidWorksApplication == null)
            {
                throw new ArgumentNullException(nameof(SolidWorksDocumentProvider._solidWorksApplication), "SolidWorks application is required.");
            }
            if (SolidWorksDocumentProvider.GetActiveDoc() == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the model doc object
            _tankSiteModelDoc = SolidWorksDocumentProvider.GetActiveDoc();

            // Create a default instance of the TankSiteAssemblySettings class
            // to hold settings for the TankSiteAssembly object.
            _tankSiteAssemblySettings = new TankSiteAssemblySettings();
            _compartmentsManager = new CompartmentsManager();
            _assemblyOfDishedEnds = new AssemblyOfDishedEnds();
            _assemblyOfCylindricalShells = new AssemblyOfCylindricalShells();

            try
            {
                _tankSiteAssemblySettings.AddTankSiteAssemblyPersistentReferenceIds(_tankSiteModelDoc);
                _compartmentsManager = _tankSiteAssemblySettings.AddCompartmentsManagerPIDs(_tankSiteModelDoc);
                _assemblyOfDishedEnds = _tankSiteAssemblySettings.AddDishedEndsPIDs(_tankSiteModelDoc);
                _assemblyOfCylindricalShells = _tankSiteAssemblySettings.AddCylindricalShellsPIDs(_tankSiteModelDoc);

                // Serialize Settings and Create Attribute
                var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
                string tankSiteAssemblyString = JsonConvert.SerializeObject(this, Formatting.Indented, options);

                AttributeManager.CreateAttribute(
                        SolidWorksDocumentProvider._solidWorksApplication,
                        _tankSiteModelDoc,
                        "TankSiteAssembly",
                        "MainEntities",
                        "MainEntities",
                        tankSiteAssemblyString);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        // Properties for Main Entities
        public Feature GetCenterAxis() => 
            (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCenterAxis, out _);

        public Component2 GetWorkshopAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankWorkshopAssembly, out _);

        public Feature GetAxisMate() => 
            (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDAxisMate, out _);

        public Component2 GetTankAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankAssembly, out _);

        public Component2 GetShellAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDShellAssembly, out _);

        public Component2 GetDishedEndsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDDishedEndsAssembly, out _);

        public Component2 GetCylindricalShellsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCylindricalShellsAssembly, out _);

        public Component2 GetCompartmentsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCompartmentsAssembly, out _);

    }
}

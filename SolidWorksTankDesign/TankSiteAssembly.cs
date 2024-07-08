using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using WarningAndErrorService;
using Attribute = SolidWorks.Interop.sldworks.Attribute;

namespace SolidWorksTankDesign
{
    /* This is the highest class in the hierarchy. At least, it must have the center axis and 
     Tank component.*/
    internal class TankSiteAssembly
    {
        [JsonIgnore]
        private ModelDoc2 _tankSiteModelDoc;
        [JsonIgnore]
        private ModelDocExtension _tankSiteDocExtension => _tankSiteModelDoc.Extension;

        public TankSiteAssemblySettings _tankSiteAssemblySettings;
        public AssemblyOfDishedEnds _assemblyOfDishedEnds;

        /// <summary>
        /// This constructor is used when Json string is deserialized into an object
        /// </summary>
        public TankSiteAssembly()
        {
            // Null checks
            if (SolidWorksDocumentProvider._solidWorksApplication == null)
            {
                throw new ArgumentNullException(nameof(SolidWorksDocumentProvider._solidWorksApplication), "SolidWorks application is required.");
            }
            if (SolidWorksDocumentProvider.ActiveDoc() == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the model doc object
            _tankSiteModelDoc = SolidWorksDocumentProvider.ActiveDoc();

            // Create a default instance of the TankSiteAssemblySettings class
            // to hold settings for the TankSiteAssembly object.
            _tankSiteAssemblySettings = new TankSiteAssemblySettings();
            _assemblyOfDishedEnds = new AssemblyOfDishedEnds();
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
            if (SolidWorksDocumentProvider.ActiveDoc() == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the model doc object
            _tankSiteModelDoc = SolidWorksDocumentProvider.ActiveDoc();

            // Create a default instance of the TankSiteAssemblySettings class
            // to hold settings for the TankSiteAssembly object.
            _tankSiteAssemblySettings = new TankSiteAssemblySettings();
            _assemblyOfDishedEnds = new AssemblyOfDishedEnds();

            try
            {
                _tankSiteAssemblySettings.AddTankSiteAssemblyPersistentReferenceIds(_tankSiteModelDoc);
                _assemblyOfDishedEnds = _tankSiteAssemblySettings.AddDishedEndsPIDs(SolidWorksDocumentProvider._solidWorksApplication, _tankSiteModelDoc);

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
        public Feature centerAxis() => 
            (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCenterAxis, out _);

        public Component2 workshopAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankWorkshopAssembly, out _);

        public Feature axisMate() => 
            (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDAxisMate, out _);

        public Component2 tankAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankAssembly, out _);

        public Component2 shellAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDShellAssembly, out _);

        public Component2 dishedEndsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDDishedEndsAssembly, out _);

        public Component2 cylindricalShellsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCylindricalShellsAssembly, out _);

        public Component2 compartmentsAssemblyComponent() => 
            (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCompartmentsAssembly, out _);

    }
}

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
        // Private Fields
        [JsonIgnore]
        SldWorks _solidWorksApplication;

        [JsonIgnore]
        private readonly ModelDoc2 _tankSiteModelDoc;
        [JsonIgnore]
        private ModelDocExtension _tankSiteDocExtension;

        public TankSiteAssemblySettings _tankSiteAssemblySettings;
        public AssemblyOfDishedEnds _assemblyOfDishedEnds;

        /// <summary>
        /// This constructor is used when Json string is deserialized into an object
        /// </summary>
        public TankSiteAssembly(){ }

        /// <summary>
        /// Initializes a TankSiteAssembly object, representing a SolidWorks tank site assembly model.It ensures 
        /// the provided SolidWorks application and model document are valid and performs initial setup operations.
        /// </summary>
        /// <param name="warningService"></param>
        /// <param name="solidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public TankSiteAssembly(SldWorks solidWorksApplication, ModelDoc2 tankSiteModelDoc, bool attributeExists)
        {
            // Null checks
            if (solidWorksApplication == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._solidWorksApplication), "SolidWorks application is required.");
            }

            if (tankSiteModelDoc == null)
            {
                throw new ArgumentNullException(nameof(TankSiteAssembly._tankSiteModelDoc), "Tank site model document is required.");
            }

            // Store references to the SolidWorks application and model objects
            this._solidWorksApplication = solidWorksApplication;
            this._tankSiteModelDoc = tankSiteModelDoc;
            _tankSiteDocExtension = _tankSiteModelDoc.Extension;

            // Create a default instance of the TankSiteAssemblySettings class
            // to hold settings for the TankSiteAssembly object.
            _tankSiteAssemblySettings = new TankSiteAssemblySettings();
            _assemblyOfDishedEnds = new AssemblyOfDishedEnds();

            // If the attribite already exists in the document, assign settings, 
            // else - add PIDs, serialize object and create an attribute.
            if (attributeExists)
            {
                _tankSiteAssemblySettings = LoadTankSiteAssemblySettingsFromAttribute();

                CheckForNullPropertiesAndNotify();
            }
            else
            {
                try
                {
                    _tankSiteAssemblySettings.AddTankSiteAssemblyPersistentReferenceIds(tankSiteModelDoc);
                    _assemblyOfDishedEnds = _tankSiteAssemblySettings.AddDishedEndsPIDs(solidWorksApplication, tankSiteModelDoc);

                    // Serialize Settings and Create Attribute
                    var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
                    string tankSiteAssemblyString = JsonConvert.SerializeObject(this, Formatting.Indented, options);

                    AttributeManager.CreateAttribute(
                            solidWorksApplication,
                            tankSiteModelDoc,
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

            /// <summary>
            /// Initializes and assigns settings for the TankSiteAssembly object.
            /// Attempts to retrieve settings from a SolidWorks attribute and deserialize them from JSON.
            /// If retrieval or deserialization fails, a default TankSiteAssemblySettings object is used.
            /// </summary>
            TankSiteAssemblySettings LoadTankSiteAssemblySettingsFromAttribute()
            {
                // 1. Attempt to retrieve the parameter value from the SolidWorks document.
                // The value is expected to be a JSON string containing the serialized settings.
                string parameterValue = null;
                try
                {
                    parameterValue = AttributeManager.GetAttributeParameterValue(tankSiteModelDoc, "MainEntities", "MainEntities");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + " Tank site assembly settings could not be set.");
                    return null;
                }

                try
                {
                    // Create a JsonSerializerSettings object to configure deserialization behavior.
                    var Jsonsettings = new JsonSerializerSettings
                    {
                        // This setting tells the deserializer to replace existing object properties with new values.
                        ObjectCreationHandling = ObjectCreationHandling.Replace,
                        // This contract resolver enables deserialization of private properties.
                        ContractResolver = new PrivatePropertyContractResolver()
                    };

                    // Deserialize the JSON into a temporary object with a structure matching the JSON
                    // This allows us to extract the nested '_tankSiteAssemblySettings' object later
                    var deserializedObject = JsonConvert.DeserializeAnonymousType
                    (parameterValue,
                    new
                    {
                        _tankSiteAssemblySettings = new TankSiteAssemblySettings() // Empty object to be populated
                    },
                    Jsonsettings);

                    // Extract the TankSiteAssemblySettings object from the deserialized anonymous object
                    // This is where the actual values from the JSON are assigned to our settings object
                    _tankSiteAssemblySettings = deserializedObject._tankSiteAssemblySettings; // Assign the inner object
                }

                catch (JsonException ex)
                {
                    MessageBox.Show("Error deserializing tank site assembly settings: " + ex.Message);
                    return null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while assigning settings: " + ex.Message);
                    return null;
                }

                return _tankSiteAssemblySettings;
            }

            /// <summary>
            /// Checks the TankSiteAssemblySettings object for any properties with null values
            /// and displays a message box listing them if any are found.
            /// </summary>
            void CheckForNullPropertiesAndNotify()
            {
                // 1. Get all properties of the TankSiteAssemblySettings class using reflection.
                var properties = typeof(TankSiteAssemblySettings).GetProperties();

                // 2. Filter the properties to only include those with null values in the current settings object.
                var nullProperties = properties.Where(prop => prop.GetValue(_tankSiteAssemblySettings) == null);

                // 3. Project the filtered properties into an enumerable of their names as strings.
                var nullPropertyNames = nullProperties.Select(prop => prop.Name);

                // 4. Check if there are any null properties.
                if (nullProperties.Any())
                {
                    // 5.Create a user-friendly message listing the names of the null properties.
                    string message = $"The following properties could not be set in TankSiteAssemblySettings: {string.Join(", ", nullProperties)}";

                    // 6. Display the message to the user in a message box.
                    MessageBox.Show(message);
                    return;
                }
            }
        }

        // Properties for Main Entities
        public Feature centerAxis => (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCenterAxis, out _);
        public Component2 workshopAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankWorkshopAssembly, out _);
        public Feature axisMate => (Feature)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDAxisMate, out _);
        public Component2 tankAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDTankAssembly, out _);
        public Component2 shellAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDShellAssembly, out _);
        public Component2 dishedEndsAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDDishedEndsAssembly, out _);
        public Component2 cylindricalShellsAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCylindricalShellsAssembly, out _);
        public Component2 compartmentsAssembly => (Component2)_tankSiteDocExtension.GetObjectByPersistReference3(_tankSiteAssemblySettings.PIDCompartmentsAssembly, out _);

    }
}

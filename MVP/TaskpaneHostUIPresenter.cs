using Newtonsoft.Json;
using SolidWorksTankDesign;
using System.Windows.Forms;
using System;

namespace MVP
{
    public class TaskpaneHostUIPresenter
    {
        private readonly ITaskpaneHostUI _taskpaneHostUIView;

        public TaskpaneHostUIPresenter(ITaskpaneHostUI taskpaneHostUIView)
        {
            _taskpaneHostUIView = taskpaneHostUIView;

            _taskpaneHostUIView.UpdateSettings += _taskpaneHostUIView_UpdateSettings;
            _taskpaneHostUIView.CreateTank += _taskpaneHostUIView_CreateTank;
            _taskpaneHostUIView.RecognizeTankAssemlby += _taskpaneHostUIView_RecognizeTankAssemlby;
        }

        private void _taskpaneHostUIView_RecognizeTankAssemlby(object sender, System.EventArgs e)
        {
            TankSiteAssembly tankSiteAssembly = new TankSiteAssembly();
            tankSiteAssembly.InitializeAndStoreTankSiteConfiguration();

            SolidWorksDocumentProvider._tankSiteAssembly = LoadTankSiteAssemblySettingsFromAttribute();
        }

        private void _taskpaneHostUIView_CreateTank(object sender, System.EventArgs e)
        {

        }

        private void _taskpaneHostUIView_UpdateSettings(object sender, System.EventArgs e)
        {

        }

        /// <summary>
        /// Initializes and assigns settings for the TankSiteAssembly object.
        /// Attempts to retrieve settings from a SolidWorks attribute and deserialize them from JSON.
        /// If retrieval or deserialization fails, a default TankSiteAssemblySettings object is used.
        /// </summary>
        TankSiteAssembly LoadTankSiteAssemblySettingsFromAttribute()
        {
            // 1. Attempt to retrieve the parameter value from the SolidWorks document.
            // The value is expected to be a JSON string containing the serialized settings.
            string parameterValue = null;
            TankSiteAssembly deserializedObject = null;
            try
            {
                parameterValue = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.GetActiveDoc(), "MainEntities", "MainEntities");
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
                deserializedObject = JsonConvert.DeserializeObject<TankSiteAssembly>(parameterValue, Jsonsettings);

                // Extract the TankSiteAssemblySettings object from the deserialized anonymous object
                // This is where the actual values from the JSON are assigned to our settings object
                //_tankSiteAssemblySettings = deserializedObject._tankSiteAssemblySettings;
                //_assemblyOfDishedEnds = deserializedObject._assemblyOfDishedEnds;// Assign the inner object
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

            return deserializedObject;
        }
    
    }
}

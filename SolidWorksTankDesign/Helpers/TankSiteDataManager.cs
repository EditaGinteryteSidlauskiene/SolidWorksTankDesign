using Newtonsoft.Json;
using SolidWorks.Interop.swconst;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class TankSiteDataManager
    {
        /// <summary>
        /// Initializes and assigns settings for the TankSiteAssembly object.
        /// Attempts to retrieve settings from a SolidWorks attribute and deserialize them from JSON.
        /// If retrieval or deserialization fails, a default TankSiteAssemblySettings object is used.
        /// </summary>
        public static TankSiteAssembly LoadTankSiteAssemblyFromAttribute()
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
                //_tankSiteAssemblySettings = filesToDeleteList._tankSiteAssemblySettings;
                //_assemblyOfDishedEnds = filesToDeleteList._assemblyOfDishedEnds;// Assign the inner object
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

        /// <summary>
        /// Initializes and assigns settings for the TankSiteAssembly object.
        /// Attempts to retrieve settings from a SolidWorks attribute and deserialize them from JSON.
        /// If retrieval or deserialization fails, a default TankSiteAssemblySettings object is used.
        /// </summary>
        public static TankProperties LoadTankSiteAssemblyPropertiesFromAttribute()
        {
            // 1. Attempt to retrieve the parameter value from the SolidWorks document.
            // The value is expected to be a JSON string containing the serialized settings.
            string parameterValue = null;
            TankProperties deserializedObject = null;
            try
            {
                parameterValue = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.GetActiveDoc(), "MainEntities", "TankProperties");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " Tank site assembly properties could not be set.");
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
                deserializedObject = JsonConvert.DeserializeObject<TankProperties>(parameterValue, Jsonsettings);

                // Extract the TankSiteAssemblySettings object from the deserialized anonymous object
                // This is where the actual values from the JSON are assigned to our settings object
                //_tankSiteAssemblySettings = filesToDeleteList._tankSiteAssemblySettings;
                //_assemblyOfDishedEnds = filesToDeleteList._assemblyOfDishedEnds;// Assign the inner object
            }

            catch (JsonException ex)
            {
                MessageBox.Show("Error deserializing tank site assembly properties: " + ex.Message);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while assigning properties: " + ex.Message);
                return null;
            }

            return deserializedObject;
        }

        public static List<string> LoadFilesToDeleteFromAttribute()
        {
            // 1. Attempt to retrieve the parameter value from the SolidWorks document.
            // The value is expected to be a JSON string containing the serialized settings.
            string parameterValue = null;
            List<string> filesToDeleteList = null;
            try
            {
                parameterValue = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.GetActiveDoc(), "MainEntities", "FilesToDelete");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " Files to delete could not be set.");
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
                filesToDeleteList = JsonConvert.DeserializeObject<List<string>>(parameterValue, Jsonsettings);

                // Extract the TankSiteAssemblySettings object from the deserialized anonymous object
                // This is where the actual values from the JSON are assigned to our settings object
                //_tankSiteAssemblySettings = filesToDeleteList._tankSiteAssemblySettings;
                //_assemblyOfDishedEnds = filesToDeleteList._assemblyOfDishedEnds;// Assign the inner object
            }

            catch (JsonException ex)
            {
                MessageBox.Show("Error deserializing files to delete list: " + ex.Message);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while assigning files to delete list: " + ex.Message);
                return null;
            }

            return filesToDeleteList;
        }


        /// <summary>
        /// Serializes the current TankSiteAssembly object (including private properties) into a JSON string 
        /// and stores it within a specified SolidWorks model attribute. This method is used parts or assembies are added, or deleted
        /// </summary>
        public static void SerializeAndStoreTankSiteAssemblyData()
        {
            // Configure JSON serialization to include private properties
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };

            // Serialize the entire TankSiteAssembly object into a formatted JSON string
            string tankSiteAssemblyString = JsonConvert.SerializeObject(SolidWorksDocumentProvider._tankSiteAssembly, Formatting.Indented, options);

            try
            {
                // Update the specified attribute in the SolidWorks model with the serialized data
                AttributeManager.EditAttributeParameterValue(
                    SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc,
                    "MainEntities",
                    "MainEntities",
                    tankSiteAssemblyString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to serialize and store tank site assembly data. " + ex.Message);
            }
        }

        public static void UpdateTankProperties()
        {
            // Serialize properties and update Attribute
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
            string tankPropertiesString = JsonConvert.SerializeObject(SolidWorksDocumentProvider._tankProperties, Formatting.Indented, options);

            AttributeManager.EditAttributeParameterValue(
                SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc,
                "MainEntities",
                "TankProperties",
                tankPropertiesString);

            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }

        public static void UpdateFilesToDeleteList()
        {
            // Serialize properties and update Attribute
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
            string filesToDelete = JsonConvert.SerializeObject(SolidWorksDocumentProvider._filesToDelete, Formatting.Indented, options);

            AttributeManager.EditAttributeParameterValue(
                SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc,
                "MainEntities",
                "FilesToDelete",
                filesToDelete);

            SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc.Save3(
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                (int)swFileSaveError_e.swGenericSaveError,
                (int)swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild);
        }
    }
}

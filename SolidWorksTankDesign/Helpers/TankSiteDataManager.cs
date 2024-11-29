using Newtonsoft.Json;
using System;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    internal class TankSiteDataManager
    {
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
    }
}

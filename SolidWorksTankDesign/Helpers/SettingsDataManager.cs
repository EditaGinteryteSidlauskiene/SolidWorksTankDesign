using Newtonsoft.Json;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Helpers
{
    internal static class SettingsDataManager
    {
        private const string PAINTING_SYSTEMS_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Painting Systems.txt";

        /// <summary>
        /// Gets a list of treatments from the document
        /// </summary>
        /// <returns></returns>
        public static List<Treatment> GetTreatments()
        {
            // Create a JsonSerializerSettings object to configure deserialization behavior.
            var Jsonsettings = new JsonSerializerSettings
            {
                // This setting tells the deserializer to replace existing object properties with new values.
                ObjectCreationHandling = ObjectCreationHandling.Replace,
                // This contract resolver enables deserialization of private properties.
                ContractResolver = new PrivatePropertyContractResolver()
            };

            try
            {
                // Get all lines in settings file
                string paintingSettings = File.ReadAllText(PAINTING_SYSTEMS_DOC_PATH);

                return JsonConvert.DeserializeObject<List<Treatment>>(paintingSettings, Jsonsettings);
            }
            catch (Exception ex)
            {
                // Handle exceptions
                MessageBox.Show($"Settings document was not found. Check documents path {PAINTING_SYSTEMS_DOC_PATH}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Saves changes made in treatment in settings document
        /// </summary>
        /// <param name="treatments"></param>
        public static void SaveChangesInTreatments(List<Treatment> treatments)
        {
            // Configure JSON serialization to include private properties
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };

            // Serialize the entire TankSiteAssembly object into a formatted JSON string
            string treatmentsString = JsonConvert.SerializeObject(treatments, Formatting.Indented, options);

            File.WriteAllText(PAINTING_SYSTEMS_DOC_PATH, treatmentsString);
        }
    }
}

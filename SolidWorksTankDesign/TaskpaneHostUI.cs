﻿using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        // Getting SW application reference.
        public void AssignSWApp(SldWorks inputswApp)
        {
            SolidWorksDocumentProvider._solidWorksApplication = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //TankSiteAssembly tankSiteAssembly = new TankSiteAssembly();
            //tankSiteAssembly.InitializeAndStoreTankSiteConfiguration();

            SolidWorksDocumentProvider._tankSiteAssembly = LoadTankSiteAssemblySettingsFromAttribute();

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

            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

            //tankSiteAssembly._compartmentsManager.Compartments[0].ActivateDocument();
            //tankSiteAssembly._compartmentsManager.Compartments[0].Nozzles[0].AddNozzleAssembly();

            //tankSiteAssembly._assemblyOfDishedEnds.SetNumberOfInnerDishedEnds(1, DishedEndAlignment.Left, 2);
            //tankSiteAssembly._assemblyOfCylindricalShells.SetNumberOfCylindricalShells(2, 2, 2500);

            //tankSiteAssembly._compartmentsManager.SetNumberOfCompartments(2, tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[0], 1);

            //////tankSiteAssembly._compartmentsManager.Compartments[1].ActivateDocument();
            //////tankSiteAssembly._compartmentsManager.Compartments[1].AddNozzle(2, tankSiteAssembly._compartmentsManager.Compartments[1].GetLeftEndPlane(), 1);
            //tankSiteAssembly._compartmentsManager.ActivateDocument();
            //tankSiteAssembly._compartmentsManager.SetNumberOfCompartments(1, null, 0);

            //int count = tankSiteAssembly._compartmentsManager.Compartments.Count;

            //tankSiteAssembly._compartmentsManager.Compartments[1].ActivateDocument();
            //tankSiteAssembly._compartmentsManager.Compartments[1].AddNozzle(2, tankSiteAssembly._compartmentsManager.Compartments[1].Nozzles[0].GetPositionPlane(), 1);
        }
    }
}

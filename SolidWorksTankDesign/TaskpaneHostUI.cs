using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Xml.Linq;
using WarningAndErrorService;
using Attribute = SolidWorks.Interop.sldworks.Attribute;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Numerics;
using System.Linq;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private WarningService warningService = new WarningService();

        // Getting SW application reference.
        public void AssignSWApp(SldWorks inputswApp)
        {
            SolidWorksDocumentProvider._solidWorksApplication = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ModelDoc2 tankSiteModelDoc = SolidWorksDocumentProvider._solidWorksApplication.IActiveDoc2;

            TankSiteAssembly tankSiteAssembly = LoadTankSiteAssemblySettingsFromAttribute();

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
                    parameterValue = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.ActiveDoc(), "MainEntities", "MainEntities");
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


            ModelDoc2 doc = tankSiteAssembly.dishedEndsAssemblyComponent().GetModelDoc2();
            SolidWorksDocumentProvider._solidWorksApplication.ActivateDoc3(doc.GetTitle(), true, 0, 0);

            tankSiteAssembly._assemblyOfDishedEnds.AddInnerDishedEnd(
                tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd,
                DishedEndAlignment.Left,
                2, 1);

            byte[] n = tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[0]._dishedEndSettings.PIDRightPlaneMate;
        }
    }
}

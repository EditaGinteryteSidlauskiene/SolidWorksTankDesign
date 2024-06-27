using AddinWithTaskpane;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using WarningAndErrorService;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private WarningService warningService = new WarningService();

        // Getting SW application reference.
        SldWorks swApp;
        public void AssignSWApp(SldWorks inputswApp)
        {
            swApp = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            ModelDoc2 tankSiteModelDoc = swApp.IActiveDoc2;

            // ------------------ Create attribute ------------------------

            // Create attribute definition for attribute of tank component
            AttributeDef attributeDefinitionTankComponent = swApp.DefineAttribute("TankComponent");

            // Add parameter for attribute Type - name of parameter, 1 - value (attribute of tank component)
            attributeDefinitionTankComponent.AddParameter(
                "Type",
                (int)swParamType_e.swParamTypeDouble,
                1,
                0);

            // Register added parameter to attribute definition
            bool isRegistered = attributeDefinitionTankComponent.Register();

            // Get tank as component
            Component2 tankComponent = Utilities.GetTopLevelComponents(tankSiteModelDoc)[0];

            // Create invisible attribute for tank component in TankSiteAssembly document
            Attribute attributeTankComponent = attributeDefinitionTankComponent.CreateInstance5(
                tankSiteModelDoc,
                tankComponent,
                "TankComponent",
                1,
                (int)swInConfigurationOpts_e.swThisConfiguration);

            //------------------------ Attribute is created --------------------------------------

            //------------------------- Get tank component from attribute -----------------------------

            // Get first attribute as feature
            Feature attributeFeature = Utilities.GetNthFeatureOfType(
                tankSiteModelDoc, 
                FeatureType.Attribute, 
                1);

            // Get feature's attribute
            Attribute attributeToCheck = attributeFeature.GetSpecificFeature2();

            // Get parameter with name "Type" of the attribute
            Parameter parameterToCheck = attributeToCheck.GetParameter("Type");

            // Check is the value of the parameter = 1
            if(parameterToCheck.GetDoubleValue() == 1)
            {
                // Get component which the attribute belongs to
                Component2 tankAsComponent = attributeToCheck.GetComponent();

                //Get component's (Tank) component (Shell)
                Component2 shellAsComponent = Utilities.GetAllComponents(tankAsComponent)[0].Component;

                //----------------- Create attribute for shell as component ------------------------

                // Create attribute definition for shell component
                AttributeDef attributeDefinitionShellComponent = swApp.DefineAttribute("ShellComponent");

                // Add parameter with name "Type", and value = 2
                attributeDefinitionShellComponent.AddParameter(
                    "Type",
                    (int)swParamType_e.swParamTypeDouble,
                    2,
                    0);

                // Register this parameter into attribute definition
                attributeDefinitionShellComponent.Register();

                // Create attribute for shell component in the TanksSiteAssembly document.
                Attribute attributeShellComponent = attributeDefinitionShellComponent.CreateInstance5(
               tankSiteModelDoc,
               shellAsComponent,
               "ShellComponent",
               1,
               (int)swInConfigurationOpts_e.swThisConfiguration);

                //---------------------- Attribute is created ---------------------------
            }

            //------------------------ Access component through attribute -----------------------------

            // Get second attribute as feature
            attributeFeature = Utilities.GetNthFeatureOfType(
                tankSiteModelDoc,
                FeatureType.Attribute,
                2);

            // Get attribute of this feature
            attributeToCheck = attributeFeature.GetSpecificFeature2();

            // Get parameter of the attribute
            parameterToCheck = attributeToCheck.GetParameter("Type");

            // Check if the value of the parameter = 2
            if (parameterToCheck.GetDoubleValue() == 2)
            {
                // Get component whic the attribute belongs to
                Component2 shellAsComponent = attributeToCheck.GetComponent();

                // Select the component
                shellAsComponent.Select2(false, 1);
            }


            //TankSiteAssembly tankSiteAssembly = new TankSiteAssembly(warningService, swApp, tankSiteModelDoc);
        }
    }
}

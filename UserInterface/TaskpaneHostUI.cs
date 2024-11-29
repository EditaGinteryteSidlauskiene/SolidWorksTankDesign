using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private const string SETTINGS_FILE_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Settings.txt";

        // Getting SW application reference.
        public void AssignSWApp(SldWorks inputswApp)
        {
            SolidWorksDocumentProvider._solidWorksApplication = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent(SETTINGS_FILE_PATH);
        }

        private void RecognizeButton_Click(object sender, EventArgs e)
        {
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
                string mainEntitiesString = null;
                string tankPropertiesString = null;
                TankSiteAssembly deserializedObject = null;
                SolidWorksDocumentProvider._tankProperties = new TankProperties();
                try
                {
                    mainEntitiesString = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.GetActiveDoc(), "MainEntities", "MainEntities");
                    tankPropertiesString = AttributeManager.GetAttributeParameterValue(SolidWorksDocumentProvider.GetActiveDoc(), "MainEntities", "TankProperties"); ;
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
                    deserializedObject = JsonConvert.DeserializeObject<TankSiteAssembly>(mainEntitiesString, Jsonsettings);
                    SolidWorksDocumentProvider._tankProperties = JsonConvert.DeserializeObject<TankProperties>(tankPropertiesString, Jsonsettings);

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

        private void NewButton_Click(object sender, EventArgs e)
        {
            InitialConfiguration initialConfiguration = new InitialConfiguration(SETTINGS_FILE_PATH);

            Controls.Add(initialConfiguration);
            initialConfiguration.BringToFront();
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            SettingsWindow settingsWindow = new SettingsWindow(SETTINGS_FILE_PATH);

            Controls.Add(settingsWindow);
            settingsWindow.BringToFront();
        }

        //private void AddFirstNozzleButton_Click(object sender, EventArgs e)
        //{
        //    _tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();
        //    //_tankSiteAssembly._assemblyOfCylindricalShells.SetNumberOfCylindricalShells(1, 1.5, 2500);

        //    Compartment compartment = _tankSiteAssembly._compartmentsManager.Compartments[0];
        //    Nozzle nozzle = compartment.Nozzles[0];

        //    _tankSiteAssembly._compartmentsManager.Compartments[0].ActivateDocument();
        //    nozzle.AddNozzleAssembly(compartment, nozzle);
        //}

        //private void SetSectionTextBox_KeyUp(object sender, KeyEventArgs e)
        //{
        //    //If Enter is pressed
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        //Validate and assign values from input
        //        if (!ParseAndValidateTankConfigurationInput(e))
        //        {
        //            return;
        //        }

        //        ConfigureTankShellAndCompartments();

        //        if(compartmentConfiguration.Count>0)
        //        {
        //            AdjustNozzleCompartments();
        //        }
        //    }
        //}


        //private void AddNozzleSubmitButton_Click(object sender, EventArgs e)
        //{
        //    _tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

        //    int compartmentNumber = (int)numericUpDown1.Value;
        //    string inputLength = setSectionsTextBox.Text.Replace(',', '.');
        //    double distance = double.Parse(inputLength);

        //    _tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber - 1].ActivateDocument();
        //    Feature nozzleRefPlane = _tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber - 1].Nozzles.Last().GetPositionPlane();

        //    _tankSiteAssembly._compartmentsManager.Compartments[compartmentNumber - 1].AddNozzle(compartmentNumber, nozzleRefPlane, distance);
        //}




        ///// <summary>
        ///// Adds UI elements to set lengths of required number of cylindrical shells.
        ///// </summary>
        //private void AdjustFieldsForCylindricalShells()
        //{
        //    int numberOfCylindricalShells = (int)numericUpDown2.Value;
        //    int lastPanelHeightY = 0;
        //    int xOffset = 0; // Starting X position for controls

        //    // Clears panel where labels and input text boxes are stored
        //    manholeSettingPanel.Controls.Clear();

        //    // Keep track of the last Y position. It helps to position lines one after another
        //    int lastLocationY = 10; 

        //    for (int i = 1; i <= numberOfCylindricalShells; i++)
        //    {
        //        // Add a label
        //        Label label = new Label();
        //        label.Text = $"Enter length for cylindrical shell no {i}";
        //        label.AutoSize = true;
        //        label.Location = new System.Drawing.Point(xOffset, lastLocationY + 10 * i);
        //        manholeSettingPanel.Controls.Add(label);

        //        // Add input text box
        //        TextBox inputBox = new TextBox();
        //        inputBox.Name = "inputBox" + i; // Optional: Assign unique names
        //        inputBox.Size = Size = new System.Drawing.Size(49, 13);
        //        inputBox.Location = new System.Drawing.Point(xOffset + label.Size.Width + 10, lastLocationY + 10 * i);
        //        manholeSettingPanel.Controls.Add(inputBox);

        //        // Calculate the Y-coordinate for the next control, 
        //        // with 25 pixels of spacing between each control.
        //        lastLocationY = 10 + 25 * i;
        //    }
        //    // Iterate through the panel's controls to find the lowest one
        //    foreach (Control control in this.manholeSettingPanel.Controls)
        //    {
        //        if (control.Bottom > lastPanelHeightY)
        //        {
        //            lastPanelHeightY = control.Bottom;
        //        }
        //    }

        //    // Adjust the panel's height to accommodate its contents, add 35 padding
        //    int panelHeight = lastPanelHeightY + 35;
        //    // Add some padding
        //    manholeSettingPanel.Height = panelHeight;

        //    // Add button
        //    Button addButton = new Button();
        //    addButton.Text = "Add";
        //    addButton.Size = new System.Drawing.Size(75, 23);
        //    addButton.Location = new System.Drawing.Point(manholeSettingPanel.Right - 120, manholeSettingPanel.Height - 25);
        //    addButton.BackColor = System.Drawing.Color.White;
        //    addButton.Click += AddButton_Click;
        //    manholeSettingPanel.Controls.Add(addButton);

        //    manholeSettingPanel.Visible = true;

        //    // Refreshes the taskpane
        //    if (!SolidWorksDocumentProvider._solidWorksApplication.TaskPaneIsPinned) SolidWorksDocumentProvider._solidWorksApplication.TaskPaneIsPinned = true;
        //    else SolidWorksDocumentProvider._solidWorksApplication.TaskPaneIsPinned = false;
        //}

        //private void numberOfCylindricalShells_ValueChanged(object sender, EventArgs e) => AdjustFieldsForCylindricalShells();

        ///// <summary>
        ///// Adds cylindrical shells lengths into a list
        ///// </summary>
        ///// <returns></returns>
        //private List<double> GetListOfCylindricalShellsLengths()
        //{
        //    List<double> lengths = new List<double>();

        //    // Loop each control in manholeSettingPanel, if it is a textbox, add its value into the list.
        //    foreach (Control control in manholeSettingPanel.Controls)
        //    {
        //        if (control is TextBox textBox)
        //        {
        //            // Replace "," into "." to be able to parse into double
        //            string inputText = textBox.Text.Replace(',', '.');
        //            double.TryParse(inputText, out double cylindricalShellLength);
        //            lengths.Add(cylindricalShellLength);
        //        }
        //    }

        //    return lengths;
        //}

        ///// <summary>
        ///// Sets correct number of cylindrical shells, adjust their diameters according to the first one.
        ///// If first cylindrical shell's diameter was not found, default diameter is set to 2500 mm.
        ///// </summary>
        ///// <param name="requiredNumberOfCylindricalShells"></param>
        //private void SetCorrectNumberOfCylindricalShells(int requiredNumberOfCylindricalShells)
        //{
        //    // Get first cylindrical shell's diameter
        //    double firstCylindricalShellDiameter = _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[0].GetCylindricalShellDiameter();
        //    if (firstCylindricalShellDiameter == 0)
        //    {
        //        MessageBox.Show("Unable to get first cylindrical shell's diameter. Please change diameters manually.");
        //        firstCylindricalShellDiameter = 2500;
        //    }

        //    // Set number of cylindrical shells
        //    _tankSiteAssembly._assemblyOfCylindricalShells.SetNumberOfCylindricalShells(
        //        requiredNumberOfCylindricalShells,
        //        1,
        //        firstCylindricalShellDiameter);
        //}

        //private void AdjustNozzleCuts()
        //{
        //    foreach (Compartment compartment in _tankSiteAssembly._compartmentsManager.Compartments)
        //    {
        //        foreach (Nozzle nozzle in compartment.Nozzles)
        //        {
        //            _tankSiteAssembly._compartmentsManager.ActivateDocument();

        //            nozzle.DeleteCutExtrude();

        //            nozzle.AddCutOutExtrude(compartment, nozzle);
        //            //SolidWorksDocumentProvider.GetActiveDoc().EditRebuild3();
        //        }
        //    }

        //    _tankSiteAssembly._compartmentsManager.CloseDocument();
        //}

        ///// <summary>
        ///// Sets the required number of cylindrical shells, adjust their diameters and lengths.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void AddButton_Click(object sender, EventArgs e)
        //{
        //    _tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

        //    // Store all lengths into the list
        //    List<double> lengths = GetListOfCylindricalShellsLengths();

        //    // Set the correct number of cylindrical shells
        //    int requiredNumberOfCylindricalShells = lengths.Count;
        //    if (requiredNumberOfCylindricalShells != _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells.Count)
        //    {
        //        SetCorrectNumberOfCylindricalShells(requiredNumberOfCylindricalShells);
        //    }

        //    // Change lengths of cylindrical shells
        //    for (int i = 0; i < requiredNumberOfCylindricalShells; i++)
        //    {
        //        double length = 0;
        //        if (lengths[i] == 0)
        //            length = 1;

        //        else length = lengths[i];

        //        _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[i].ChangeLength(length);
        //    }

        //    AdjustNozzleCuts();
        //}

        //private void NumberOfCylindricalShells_KeyUp(object sender, KeyEventArgs e)
        //{
        //    if (e.KeyCode == Keys.Enter)
        //        AdjustFieldsForCylindricalShells();
        //}


    }
}

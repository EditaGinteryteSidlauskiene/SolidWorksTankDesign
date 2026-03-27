using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Windows
{
    public partial class InitialConfiguration : UserControl
    {
        private const string DISHED_END_MAIN_DOC_NAME = "Main dished end.SLDPRT";
        private const string INNER_DISHED_END_DOC_NAME = "Inner dished end.SLDPRT";
        private const string CYLINDRICAL_SHELL_DOC_NAME = "Cylindrical shell.SLDPRT";
        private const string NOZZLE_POSITION_SKETCH_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Nozzle position sketch.SLDASM";
        private const string MAIN_FOLDER_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite";
        private const string EMPTY_TANKSITE_ASSEMBLY_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Tank Site Assembly.SLDASM";
        private const string COMPARTMENT_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Compartment.SLDASM";
        private const string DISHED_END_VOLUME_INCLUDE_PROP_NAME = "volume include";
        private const string DISHED_END_VOLUME_EXCLUDE_PROP_NAME = "volume exclude";
        private const string DISHED_END_AREA_INTERNAL_PROP_NAME = "area internal";
        private const string DISHED_END_AREA_EXTERNAL_PROP_NAME = "area external";
        private const string CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME = "volume per meter";
        private const double INITIAL_COMPARTMENT_LENGTH = 2000;

        private string _settingsFilePath;
        private string _folderForAllProjects;
        private string _projectFolder;

        private double _mainDishedEndIncludeVolume;
        private double _innerDishedEndExcludeVolume;
        private double _innerDishedEndIncludeVolume;
        private double _cylindricalShellVolumePerMeter;
        private List<Treatment> _externalTreatments = new List<Treatment>();

        private TankSiteAssembly _tankSiteAssembly;

        string instalation;
        double diameter;
        double initialCompartmentVolume;

        List<(DishedEndAlignment Alignment, double Volume)> compartmentConfiguration =
            new List<(DishedEndAlignment Alignment, double Volume)>();

        /// <summary>
        /// Constructor to initialize the InitialConfiguration window
        /// </summary>
        public InitialConfiguration(string settingsFilePath)
        {
            _externalTreatments = GetExternalTreatments(1);
            InitializeComponent(_externalTreatments);

            _settingsFilePath = settingsFilePath;
            GetProjectsFolder();
        }

        /// <summary>
        /// Gets a list of external treatments. Values for treatmentExternalType:
        /// 1 - for ExternalUnderGround
        /// 2 - for ExternalAboveGround
        /// </summary>
        /// <param name="treatmentExternalType"></param>
        private List<Treatment> GetExternalTreatments(int treatmentExternalType)
        {

            return (List<Treatment>)SettingsDataManager.GetTreatments();
                //.
                //Where(t => t.Type == (Treatment.TreatmentType)treatmentExternalType);
        }

        /// <summary>
        /// Sets the folder where to store this project's folder.
        /// Default folder's path is stored in Settings.txt
        /// </summary>
        private void GetProjectsFolder()
        {
            using (StreamReader reader = new StreamReader(_settingsFilePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (line.StartsWith("DefaultProjectFolder="))
                    {
                        _folderForAllProjects = line.Substring("DefaultProjectFolder=".Length);
                        _folderForAllProjects = _folderForAllProjects.Replace("\"", string.Empty);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Sets and saves a default folder of projects in the Settings document.
        /// </summary>
        private void SetProjectsFolder()
        {
            // Get all lines from Settings document
            string[] lines = File.ReadAllLines(_settingsFilePath);

            // Search for DefaultProjectFolder line
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("DefaultProjectFolder="))
                {
                    // Change default projects folder path into a new one
                    lines[i] = "DefaultProjectFolder=" + _folderForAllProjects;
                    break;
                }
            }
            // Save changes in the Settings document
            try
            {
                File.WriteAllLines(_settingsFilePath, lines);
            }
            // Catch an error
            catch (Exception ex)
            {
                if (ex is IOException)
                {
                    MessageBox.Show("The settings file is currently in use. " +
                                "Please close any programs that might be accessing it " +
                                "and try again.", "Error", MessageBoxButtons.OK);
                }
            }
            // Check default folder radio button
            finally
            {
                defaultFolderRadioButton.Checked = true;
            }
        }

        /// <summary>
        /// Copies the empty tank site assembly document and documents of components that are in tank site assembly.
        /// Documents are copied and renamed to the folder selected by the user
        /// </summary>
        /// <param name="solidWorksApp"></param>
        /// <param name="emptyTankSiteAssemblyDoc"></param>
        private string CopyDocuments(SldWorks solidWorksApp, ModelDoc2 emptyTankSiteAssemblyDoc, string serialNumber)
        {
            // Packs all documents that are in emptyTankSiteAssemblyDoc and saves them in a new foler
            string path = DocumentManager.PackAndGo(_folderForAllProjects, emptyTankSiteAssemblyDoc, null, serialNumber);

            // Get path of folder, where all documents of the current project will be stored
            string documentDirectory = Path.GetDirectoryName(path);

            // Open just copied and renamed tank site assembly document
            DocumentSpecification documentSpecification = (DocumentSpecification)solidWorksApp.GetOpenDocSpec(path);
            solidWorksApp.OpenDoc7(documentSpecification);

            // Close the primary empty tank site assembly doc
            solidWorksApp.CloseDoc(emptyTankSiteAssemblyDoc.GetTitle());

            return documentDirectory;
        }

        private void SerialNumberTextBox_GotFocus(object sender, EventArgs e)
        {
            serialNumberTextBox.Text = string.Empty;
            serialNumberTextBox.ForeColor = Color.Black;
        }

        private bool IsSerialNumberInputValid()
        {
            bool result = true;

            // Check if serial number is entered
            if (serialNumberTextBox.Text.Trim() == string.Empty)
            {
                serialNumberTextBox.Text = "Enter No.";
                serialNumberTextBox.ForeColor = Color.DarkRed;

                result = false;
            }
            // Check if there are no forbidden characters in serial no. These characters cannot be used in folder's name
            else if (serialNumberTextBox.Text.Contains('<') || serialNumberTextBox.Text.Contains('>') || serialNumberTextBox.Text.Contains(':') ||
                serialNumberTextBox.Text.Contains('"') || serialNumberTextBox.Text.Contains('/') || serialNumberTextBox.Text.Contains('\\') ||
                serialNumberTextBox.Text.Contains('|') || serialNumberTextBox.Text.Contains('?') || serialNumberTextBox.Text.Contains('*'))
            {
                MessageBox.Show("Characters that are not allowed in Serial no.:\n <   >   :   \"   /   \\   |   ?   *");
                serialNumberTextBox.Text = string.Empty;

                result = false;
            }
            // Check if serial no. is not a reserved name for folders
            else if (serialNumberTextBox.Text == "CON" || serialNumberTextBox.Text == "PRN" || serialNumberTextBox.Text == "AUX" ||
                serialNumberTextBox.Text == "NUL" || serialNumberTextBox.Text == "COM1" || serialNumberTextBox.Text == "COM2" ||
                serialNumberTextBox.Text == "COM3" || serialNumberTextBox.Text == "COM4" || serialNumberTextBox.Text == "COM5" ||
                serialNumberTextBox.Text == "COM6" || serialNumberTextBox.Text == "COM7" || serialNumberTextBox.Text == "COM8" ||
                serialNumberTextBox.Text == "COM9" || serialNumberTextBox.Text == "LPT1" || serialNumberTextBox.Text == "LPT2" ||
                serialNumberTextBox.Text == "LPT3" || serialNumberTextBox.Text == "LPT4" || serialNumberTextBox.Text == "LPT5" ||
                serialNumberTextBox.Text == "LPT6" || serialNumberTextBox.Text == "LPT7" || serialNumberTextBox.Text == "LPT8" ||
                serialNumberTextBox.Text == "LPT9")
            {
                MessageBox.Show("Serial no. cannot be:\n \"CON\", \"PRN\", \"AUX\", \"NUL\", \"COM1\", \"COM2\", \"COM3\", \"COM4\", \"COM5\"" +
                    ", \"COM6\", \"COM7\", \"COM8\", \"COM9\", \"LPT1\", \"LPT2\", \"LPT3\", \"LPT4\", \"LPT5\", \"LPT6\", \"LPT7\", \"LPT8\"" +
                    ", \"LPT9\"");
                serialNumberTextBox.Text = string.Empty;

                result = false;
            }
            // Check if serial no does not start or end with ".". This is not allowed in folder's name.
            else if (serialNumberTextBox.Text.StartsWith(".") || serialNumberTextBox.Text.EndsWith("."))
            {
                MessageBox.Show("Serial no. cannot start or end with \".\"");
                serialNumberTextBox.Text = string.Empty;

                result = false;
            }

            return result;
        }

        private bool IsMinOperatingTemperatureInputValid()
        {
            bool result = true;

            string minOperatingTempInput = MinOperatingTempTextBox.Text.Trim();

            // Check if the input is not empty
            if (minOperatingTempInput == string.Empty)
            {
                MessageBox.Show("Enter Min operating temperature.");
                result = false;
            }
            else if(!double.TryParse(minOperatingTempInput.Replace(',', '.'), out _))
            {
                MessageBox.Show("Min operating temperature must be a number.");
                result = false;
            }

            return result;
        }

        private bool IsMaxOperatingTemperatureInputValid()
        {
            bool result = true;

            string maxOperatingTempInput = MaxOperatingTempTextBox.Text.Trim();

            // Check if the input is not empty
            if (maxOperatingTempInput == string.Empty)
            {
                MessageBox.Show("Enter Max operating temperature.");
                result = false;
            }
            else if (!double.TryParse(maxOperatingTempInput.Replace(',', '.'), out _))
            {
                MessageBox.Show("Max operating temperature must be a number.");
                result = false;
            }

            return result;
        }

        private bool IsMaxOperatingPressureInputValid()
        {
            bool result = true;

            string maxOperatingPressureInput = MaxOepratingPressureTextBox.Text.Trim();

            // Check if the input is not empty
            if (maxOperatingPressureInput == string.Empty)
            {
                MessageBox.Show("Enter Max operating pressure.");
                result = false;
            }
            else if (!double.TryParse(maxOperatingPressureInput.Replace(',', '.'), out _))
            {
                MessageBox.Show("Max operating pressure must be a number.");
                result = false;
            }

            return result;
        }

        private bool IsShellLeakTestPressureInputValid()
        {
            bool result = true;

            string shellLeakTestPressureInput = ShellLeakTestPressureTextBox.Text.Trim();

            // Check if the input is not empty
            if (shellLeakTestPressureInput == string.Empty)
            {
                MessageBox.Show("Enter Shell leak test pressure.");
                result = false;
            }
            else if (!double.TryParse(shellLeakTestPressureInput.Replace(',', '.'), out _))
            {
                MessageBox.Show("Shell leak test pressure must be a number.");
                result = false;
            }

            return result;
        }

        private bool IsInterstitialSpaceLeakTestPressureInputValid()
        {
            bool result = true;

            string interstitialSpaceLeakTestPressureInput = InterstitialSpaceLeakTestPressureTextBox.Text.Trim();

            // Check if the input is not empty
            if (interstitialSpaceLeakTestPressureInput == string.Empty)
            {
                MessageBox.Show("Enter Interstitial space leak test pressure.");
                result = false;
            }
            else if (!double.TryParse(interstitialSpaceLeakTestPressureInput.Replace(',', '.'), out _))
            {
                MessageBox.Show("Interstitial space leak test pressure must be a number.");
                result = false;
            }

            return result;
        }

        private bool IsLeakDetectionSystemInputValid()
        {
            bool result = true;

            string leakDetectionSystemInput = LeakDetectionSystemTextBox.Text.Trim();

            if (leakDetectionSystemInput == string.Empty)
            {
                MessageBox.Show("Enter Leak detection system.");
                result = false;
            }

            return result;
        }

        private bool IsDiameterInputValid()
        {
            bool result = true;

            string diameterInput = DiameterTextBox.Text.Trim();

            // Check if the input is not empty
            if (diameterInput == string.Empty)
            {
                MessageBox.Show("Enter Diameter.");
                result = false;
            }
            else if (!double.TryParse(diameterInput.Replace(',', '.'), out double diameter))
            {
                MessageBox.Show("Diameter must be a number.");
                result = false;
            }
            else if(diameter <= 0)
            {
                MessageBox.Show("Invalid value for diameter.");
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Checks if the user has entered all information
        /// </summary>
        /// <returns></returns>
        private bool CheckUserInputsAndGiveMessages()
        {
            bool result = true;

            if(DoubleSkinParametersPanel.Visible == true)
            {
                result = IsLeakDetectionSystemInputValid();
            }

            return 
                IsSerialNumberInputValid() && 
                IsMinOperatingTemperatureInputValid() &&
                IsMaxOperatingTemperatureInputValid() &&
                IsMaxOperatingPressureInputValid() &&
                IsShellLeakTestPressureInputValid() &&
                IsInterstitialSpaceLeakTestPressureInputValid() &&
                IsDiameterInputValid() &&
                result;
        }

        private void DisableControls()
        {
            foreach(Control control in this.Controls) control.Enabled = false;
        }

        private void SaveTankPropertiesToAttribute()
        {
            SolidWorksDocumentProvider._tankProperties = new TankProperties();
            TankProperties tankProperties = SolidWorksDocumentProvider._tankProperties;

            tankProperties.SerialNumber = serialNumberTextBox.Text;
            tankProperties.ConstructionStandard = ConstructionStandardComboBox.SelectedItem.ToString();
            tankProperties.NominalDiameter = diameter;
            tankProperties.Class = classComboBox.SelectedItem.ToString();
            tankProperties.Type = TypeComboBox.SelectedItem.ToString();

            double.TryParse(MinOperatingTempTextBox.Text.Trim().Replace(',', '.'), out double minOperatingTemp);
            tankProperties.MinOperatingTemperature = minOperatingTemp;

            double.TryParse(MaxOperatingTempTextBox.Text.Trim().Replace(',', '.'), out double maxOperatingTemp);
            tankProperties.MaxOperatingTemperature = maxOperatingTemp;

            double.TryParse(MaxOepratingPressureTextBox.Text.Trim().Replace(',', '.'), out double maxOperatingPressure);
            tankProperties.MaxOperatingPressure = maxOperatingPressure;

            double.TryParse(ShellLeakTestPressureTextBox.Text.Trim().Replace(',', '.'), out double shellLeakTestPressure);
            tankProperties.ShellLeakTestPressure = shellLeakTestPressure;

            double.TryParse(InterstitialSpaceLeakTestPressureTextBox.Text.Trim().Replace(',', '.'), out double interstitialSpaceLeakTestPressure);
            tankProperties.InterstitialSpaceLeakTestPressure = interstitialSpaceLeakTestPressure;

            if(DoubleSkinParametersPanel.Visible == true)
            {
                tankProperties.InterstitialSpace = InterstitialSpaceComboBox.SelectedText;
                tankProperties.LeakDetectionSystem = LeakDetectionSystemTextBox.Text;
            }

            tankProperties.CompartmentsConfigurations.Add(new CompartmentConfiguration
            {
                Volume = initialCompartmentVolume,
                Length = INITIAL_COMPARTMENT_LENGTH,
                LeftDishedEndAlignment = DishedEndAlignment.Left
            });

            // Serialize properties and update Attribute
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
            string tankPropertiesString = JsonConvert.SerializeObject(tankProperties, Formatting.Indented, options);

            AttributeManager.EditAttributeParameterValue(
                SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc,
                "MainEntities",
                "TankProperties",
                tankPropertiesString);
        }

        private void ConstructionStandardComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;

            if (comboBox.SelectedIndex == 0)
                _externalTreatments =  GetExternalTreatments(1);
            else
                _externalTreatments = GetExternalTreatments(2);
        }

        private void GetInstalation()
        {
            if (ConstructionStandardComboBox.SelectedIndex == 0)
                instalation = "u";
            else
                instalation = "a";
        }

        private void GetDiameter()
        {
            double.TryParse(DiameterTextBox.Text.Replace(',', '.').Replace(" ", ""), out diameter);
        }

        /// <summary>
        /// Event, when the user clicks "Create" button.
        /// Left and right dished ends are added
        /// First cylindrical shell is added
        /// First compartment and its nozzle(s) are added
        /// Inner dished ends, more compartment and their nozzles are set
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ForwardArrow_Click(object sender, EventArgs e)
        {
            if (!CheckUserInputsAndGiveMessages())
                return;

            GetInstalation();
            GetDiameter();

            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;

            DisableControls();

            try
            {
                // Open empty tank site asssembly doc
                DocumentSpecification docSpecification =
                    (DocumentSpecification)solidWorksApp.GetOpenDocSpec(EMPTY_TANKSITE_ASSEMBLY_DOC_PATH);
                ModelDoc2 emptyTankSiteAssemblyDoc = solidWorksApp.OpenDoc7(docSpecification);

                // Get serial number, class, and number of compartments
                string serialNumber = serialNumberTextBox.Text;
                string selectedClass = classComboBox.SelectedItem.ToString();

                _projectFolder = CopyDocuments(solidWorksApp, emptyTankSiteAssemblyDoc, serialNumber);

                // Set global project folder for the running session so other parts can access it
                SolidWorksDocumentProvider.ProjectFolderPath = _projectFolder;

                // Initialize and store tank site configurations
                SolidWorksDocumentProvider._tankSiteAssembly = new TankSiteAssembly();
                SolidWorksDocumentProvider._tankSiteAssembly.InitializeAndStoreTankSiteConfiguration();

                _tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;

                // Configure tank shell and compartments according tank configuration the user entered
                ConfigureTankShellAndCompartments(COMPARTMENT_DOC_PATH);

                //compartmentConfiguration.Insert(0, (DishedEndAlignment.Left, firstVolume));
                // Close this window and open Compartments Configuration Window

                // Calculate initial compartment's volume
                //Recalculate volume
                double cylindricalShellVolume = _cylindricalShellVolumePerMeter * 2;
                initialCompartmentVolume = cylindricalShellVolume + (_mainDishedEndIncludeVolume * 2);

                CompartmentsWindow compartmentsWindow = new CompartmentsWindow(
                    MAIN_FOLDER_PATH,
                    _projectFolder,
                    NOZZLE_POSITION_SKETCH_DOC_PATH,
                    _cylindricalShellVolumePerMeter,
                    _mainDishedEndIncludeVolume,
                    _innerDishedEndIncludeVolume,
                    _innerDishedEndExcludeVolume,
                    INITIAL_COMPARTMENT_LENGTH,
                    initialCompartmentVolume);

                _tankSiteAssembly._compartmentsManager.CloseDocument();

                Controls.Add(compartmentsWindow);
                compartmentsWindow.BringToFront();

                SaveTankPropertiesToAttribute();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        ///// <summary>
        ///// This method parses and validates tank configuration input, extracting relevant data such as tank type, installation, and compartment details if the input matches the expected format. It uses regular expressions for validation and extraction, handling potential decimal separators correctly. The extracted values are then assigned to corresponding variables for further use. If the input is invalid, it displays an error message and optionally clears the input textbox. The method also prevents the default behavior of the Enter key press, which is to create a new line. 
        ///// </summary>
        ///// <param name="e"></param>
        //private bool ParseAndValidateTankConfigurationInput()
        //{
        //    //Clear the list before assigning values from user input
        //    compartmentConfiguration.Clear();

        //    // INPUT PREPARATION
        //    // ------------------------------------------------------------------
        //    // Remove spaces and convert to lowercase for case-insensitive matching
        //    string inputText = setSectionsTextBox.Text.Replace(" ", "").ToLowerInvariant();

        //    // INPUT VALIDATION
        //    // ------------------------------------------------------------------
        //    /*Use a single, descriptive regular expression to validate the entire input:
        //     - Captures:
        //       - 'type':             'd' (double) or 's' (single) for tank type
        //       - 'installation':      'a' (aboveground) or 'u' (underground) for installation type
        //       - 'diameter':         3-4 digits for tank diameter in millimeters
        //       - 'firstVolume':      First compartment's volume (number with optional decimals)
        //       - 'compartments':     Zero or more additional compartments with alignment and volume
        //     - Structure:
        //       - Tank type and installation: e.g., "da" (double, aboveground), "su" (single, underground)
        //       - Tank diameter: "t" followed by 3-4 digits (not starting with 0)
        //       - First volume: in parentheses, number with optional decimals (not starting with 0)
        //       - Compartments: 
        //          - Enclosed in parentheses
        //          - Start with '(' (left alignment) or ')' (right alignment)
        //          - Volume number (not starting with 0) with optional decimals
        //          - May be repeated for multiple compartments*/
        //    Regex regexWhole = new Regex(@"^(?'Type'[ds])(?'Instalation'[au])t(?!0)(?'Diameter'\d{3,4})\((?!0)(?'FirstVolume'\d+([,.]\d+)?)(?'Compartments'([()](?!0)\d+([,.]\d+)?)*)\)$",
        //        RegexOptions.IgnoreCase
        //        );

        //    //Is the input valid?
        //    Match match = regexWhole.Match(inputText);

        //    // SUCCESSFUL INPUT PROCESSING
        //    // ------------------------------------------------------------------
        //    if (match.Success)
        //    {
        //        // Extract captured values from the regular expression match
        //        string type = match.Groups["Type"].Value;
        //        instalation = match.Groups["Instalation"].Value;
        //        diameter = Convert.ToDouble(match.Groups["Diameter"].Value);
        //        firstVolume = Convert.ToDouble(match.Groups["FirstVolume"].Value.Replace(",", "."));

        //        // COMPARTMENT PARSING
        //        // ------------------------------------------------------------------
        //        // Extract alignment and volume data for each compartment
        //        string compartmentsInString = match.Groups["Compartments"].Value;

        //        //Create a new regex to distinguish alignment and volume from the whole string of compartments
        //        Regex regexCompartments = new Regex(@"^(?'Alignment'[()])(?!0)(?'Volume'\d+([,.]\d+)?)(?'RemainingCompartments'([()](?!0)\d+([,.]\d+)?)*)$");

        //        while (compartmentsInString != "")
        //        {
        //            //Compare regex with the string of compartments
        //            match = regexCompartments.Match(compartmentsInString);

        //            //Extract alignment and volume
        //            if (match.Success)
        //            {
        //                // Use pattern matching to determine alignment
        //                DishedEndAlignment currentAlignment = match.Groups["Alignment"].Value == "(" ? DishedEndAlignment.Left : DishedEndAlignment.Right;

        //                double currentVolume = Convert.ToDouble(match.Groups["Volume"].Value.Replace(",", "."));

        //                //Add current alignment and volume into the list
        //                compartmentConfiguration.Add((currentAlignment, currentVolume));

        //                //Take out just assigned alignment and volume from the string of compartments
        //                compartmentsInString = match.Groups["RemainingCompartments"].Value;
        //            }
        //            else
        //            {
        //                //This line should be never reached
        //                return false;
        //            }
        //        }

        //        return true;
        //    }

        //    // INPUT ERROR HANDLING
        //    // ------------------------------------------------------------------
        //    else
        //    {
        //        return false;
        //    }
        //}

        /// <summary>
        /// Gets the path for a document.
        /// Gets the folder name which depends on instalation type, diameter, and class.
        /// Gets a folder's, which is in the main folder, where empty documents are stored, path.
        /// If the folder exists, adds document's name to folder's path and returns full document path
        /// </summary>
        /// <param name="documentName"></param>
        /// <returns></returns>
        private string GetDocumentPath(string documentName)
        {
            // Get name folder where empty documents are stored
            string folderName = Path.Combine(
                $"{instalation.ToUpper()}T",
                $"{instalation.ToUpper()}T {diameter} Class {classComboBox.SelectedItem}");

            // Get the path of that folder. It is stored in the main folder
            string folderPath = Path.Combine(MAIN_FOLDER_PATH, folderName);

            // If the folder exists
            if (Directory.Exists(folderPath))
            {
                // Return the full document path
                return Path.Combine(folderPath, documentName);
            }

            // Else, message that the folder was not found
            else
            {
                MessageBox.Show($"Folder {folderName} was not found in {MAIN_FOLDER_PATH} path.");
                return string.Empty;
            }
        }

        /// <summary>
        /// Add left and right dished ends
        /// </summary>
        /// <returns></returns>
        private void AddMainDishedEnds()
        {
            // Get dished end's document's path. The document is stored in the main folder
            string dishedEndDocPath = GetDocumentPath(DISHED_END_MAIN_DOC_NAME);

            // If the dished end doc path is not empty, continue adding left and right dished ends
            if (dishedEndDocPath != string.Empty)
            {
                _tankSiteAssembly._assemblyOfDishedEnds.CompleteMainDishedEnds(_projectFolder, dishedEndDocPath);
            }

            else
                return ;
        }

        /// <summary>
        /// Add first cylindrical shell
        /// </summary>
        /// <param name="length"></param>
        private void AddFirstCylindricalShell(double length)
        {
            // Getcylindrical shell's document's path. The document is stored in the main folder
            string cylindricalShellDocPath = GetDocumentPath(CYLINDRICAL_SHELL_DOC_NAME);

            // If the cylindrical shell doc path is not empty, continue adding first cylindrical shell
            if (cylindricalShellDocPath != string.Empty)
            {
                _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells.Add(new CylindricalShell());
                _tankSiteAssembly._assemblyOfCylindricalShells._cylindricalShellDocPath = cylindricalShellDocPath;
                _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[0].CompleteFirstCylindricalShel(_projectFolder, length);
            }

            else
                return;
        }

        /// <summary>
        /// Adds first compartment
        /// </summary>
        /// <param name="projectFolder"></param>
        /// <param name="manholesSettings"></param>
        private void AddFirstCompartment(double length)
        {
            CompartmentsManager compartmentsManager = _tankSiteAssembly._compartmentsManager;
            // Activate shell document
            ModelDoc2 shellDoc = compartmentsManager.ActivateDocument();

            Feature leftDishedEndPositionPlane;

            // Get assembly of dished ends document
            ModelDoc2 assemblyOfDishedEndsDoc = _tankSiteAssembly.GetDishedEndsAssemblyComponent().GetModelDoc2();
            // Get left dished end's position plane
            using (var dishedEndsDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, assemblyOfDishedEndsDoc))
            {
                leftDishedEndPositionPlane = _tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd.GetPositionPlane();
            }

            // Add compartment
            compartmentsManager.Compartments.Add(new Compartment(
                _projectFolder,
                COMPARTMENT_DOC_PATH,
                SWFeatureManager.GetMajorPlane(shellDoc, MajorPlane.Front),
                compartmentsManager.GetCenterAxis(),
                leftDishedEndPositionPlane,
                1,
                length));

            // Get added compartment document.
            ModelDoc2 compartmentDoc = compartmentsManager.Compartments.Last().GetComponent().GetModelDoc2();
            // Get compartment's folder's directory
            string compartmentFolderDirectory = Path.GetDirectoryName(compartmentDoc.GetPathName());
        }

        /// <summary>
        /// This method is responsible for positioning the first nozzle compartment and then adjusting the total number of compartments.
        /// </summary>
        private void AdjustNozzleCompartments(
            string compartmentPath,
            List <double> compartmentsLengths)
        {
            TankSiteAssembly tankSiteAssembly = SolidWorksDocumentProvider._tankSiteAssembly;
            int compartmentsCount = compartmentConfiguration.Count;

            tankSiteAssembly._compartmentsManager.SetNumberOfCompartments(
               _projectFolder,
               compartmentPath,
               compartmentsCount + 1,
               compartmentsLengths,
               diameter);
        }

        /// <summary>
        /// Calculates the length of the compartment
        /// </summary>
        /// <param name="requiredCompartmentVolume"></param>
        /// <param name="dishedEndVolume1"></param>
        /// <param name="dishedEndVolume2"></param>
        /// <param name="cylindricalShellVolumePerMeter"></param>
        /// <returns></returns>
        private double GetCompartmentLength(double requiredCompartmentVolume, double dishedEndVolume1, double dishedEndVolume2, double cylindricalShellVolumePerMeter)
        {
            // Get volume of cylindrical shell needed
            double requiredVolumeOfCylindrcialShell = requiredCompartmentVolume - (dishedEndVolume1 + dishedEndVolume2);

            // Get length of cylindrical shell
            return requiredVolumeOfCylindrcialShell / cylindricalShellVolumePerMeter;
        }

        /// <summary>
        /// This method configures the dished ends and compartments of a tank.It sets the number of dished ends based on the provided compartment configuration, adjusts their alignment, and calculates the distances between them based on compartment volumes.It also adjusts the length of the cylindrical shell to accommodate the compartments. If there are no compartments, it simply adjusts the right dished end.
        /// </summary>
        /// <param name="TankSiteModelDoc"></param>
        private void ConfigureTankShellAndCompartments(string compartmentPath)
        {
            if (_tankSiteAssembly._compartmentsManager.Compartments.Count != 0 &&
                compartmentConfiguration.Count < _tankSiteAssembly._compartmentsManager.Compartments.Count)
            {
                //_tankSiteAssembly._compartmentsManager.SetNumberOfCompartments(
                //    null,
                //    null,
                //    null,
                //    null,
                //    compartmentConfiguration.Count + 1, 
                //    0,
                //    manholesSettings);
            }


            // DISHED ENDS AND COMPARTMENTS CONFIGURATION
            // ------------------------------------------------------------------


            // NULL REFERENCE CHECK
            // ------------------------------------------------------------------
            // Handle the scenario where _tankSiteAssembly is null due to a potential error during construction.
            if (_tankSiteAssembly == null)
            {
                MessageBox.Show("Error creating tank site assembly object.");
                return;
            }

            // Get custom properties values
            _mainDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(DISHED_END_MAIN_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _innerDishedEndExcludeVolume = -SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_EXCLUDE_PROP_NAME);
            _innerDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _cylindricalShellVolumePerMeter = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(CYLINDRICAL_SHELL_DOC_NAME), CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME);

            //// Create a list, where lengths of compartmets will be stored
            //compartmentsLengths = new List<double>();

            //// Add first compartment's length

            //// If there are no compartments
            //if (compartmentConfiguration.Count == 0)
            //    compartmentsLengths.Add(GetCompartmentLength(firstVolume, _mainDishedEndIncludeVolume, _mainDishedEndIncludeVolume, _cylindricalShellVolumePerMeter));
            //// If there is at least one compartment
            //else
            //{
            //    // Check the first compartment's dished end's (left) alignment
            //    if (compartmentConfiguration[0].Alignment == DishedEndAlignment.Left)
            //        compartmentsLengths.Add(GetCompartmentLength(
            //            firstVolume,
            //            _mainDishedEndIncludeVolume,
            //            _innerDishedEndExcludeVolume,
            //            _cylindricalShellVolumePerMeter));
            //    else
            //        compartmentsLengths.Add(GetCompartmentLength(
            //            firstVolume,
            //            _mainDishedEndIncludeVolume,
            //            _innerDishedEndIncludeVolume,
            //            _cylindricalShellVolumePerMeter));
            //}

            AddMainDishedEnds();

            AddFirstCylindricalShell(INITIAL_COMPARTMENT_LENGTH / 1000);
            AddFirstCompartment(INITIAL_COMPARTMENT_LENGTH / 1000);

            //// Sets the number of dished ends in the tank model based on the compartment configuration.
            //_tankSiteAssembly._assemblyOfDishedEnds.SetNumberOfInnerDishedEnds(
            //    _projectFolder,
            //    GetDocumentPath(INNER_DISHED_END_DOC_NAME),
            //    compartmentConfiguration.Count,
            //    DishedEndAlignment.Left,
            //    1);

            //// If there are compartments, adjust the distances and alignments of the dished ends.
            //if (compartmentConfiguration.Count > 0)
            //{
            //    // Iterates through each compartment in the configuration.
            //    for (int i = 0; i < compartmentConfiguration.Count; i++)
            //    {
            //        // INDEX OUT OF RANGE CHECK
            //        // ------------------------------------------------------------------
            //        // Ensure that 'i' is within the valid bounds of the CompartmentDishedEnds collection.
            //        if (i >= _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds.Count)
            //        {
            //            MessageBox.Show("Error adjusting compartment: index out of range.");
            //            return;
            //        }

            //        /*Adjusts the distance of the dished end (except for the first one) 
            //        based on the volume of the previous compartment.*/
            //        if (i > 0)
            //        {
            //            _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[i].ChangeDistance(compartmentsLengths[i]);
            //        }

            //        double compartmentLeftDishedEndVolume;
            //        double compartmentRightDishedEndVolume;

            //        // Get both, left and right, dished ends' volumes for each compartment

            //        // Assign include or exclude volume depending on the alignment
            //        compartmentLeftDishedEndVolume = compartmentConfiguration[i].Alignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
                    
            //        // Before assigning right dished end's volume, check if current compartment is the last or not.
            //        if (i + 1 != compartmentConfiguration.Count)
            //        {
            //            // If current compartment is not the last, assign include or exclude volume depending on the alignment
            //            compartmentRightDishedEndVolume = compartmentConfiguration[i + 1].Alignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
            //        }
            //        else
            //        {
            //            // If the current compartment is the last one, compartment's right dished end's volume is mainDishedEndIncludeVolume.
            //            compartmentRightDishedEndVolume = _mainDishedEndIncludeVolume;
            //        }

            //        // Add each compartment's length into the list
            //        compartmentsLengths.Add(GetCompartmentLength(compartmentConfiguration[i].Volume, compartmentLeftDishedEndVolume, compartmentRightDishedEndVolume, _cylindricalShellVolumePerMeter));

            //        _tankSiteAssembly._assemblyOfDishedEnds.ActivateDocument();
            //        // Sets the alignment of the current dished end based on the configuration.
            //        ComponentManager.SetAlignment(
            //            _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[i].GetComponent(),
            //            compartmentConfiguration[i].Alignment,
            //            _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[i].GetCenterAxisMate(),
            //            _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[i].GetRightPlaneMate());

            //    }

            //    // Calculate cylindrical shell's length
            //    foreach (double length in compartmentsLengths)
            //        cylindricalShellLength += length;

            //    // Adjusts the distance of the first compartment and right dished ends.
            //    _tankSiteAssembly._assemblyOfDishedEnds.InnerDishedEnds[0].ChangeDistance(compartmentsLengths[0]);
            //    _tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(compartmentsLengths.Last());
            //}

            // If there are no compartments, adjust only the distance of the right dished end.
            //else
            //{
                //cylindricalShellLength = compartmentsLengths[0];

                _tankSiteAssembly._assemblyOfDishedEnds.ActivateDocument();
                _tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(INITIAL_COMPARTMENT_LENGTH / 1000);
                _tankSiteAssembly._assemblyOfDishedEnds.CloseDocument();
            //}

            //// NULL REFERENCE CHECK
            //// ------------------------------------------------------------------
            //// Verify that the CylindricalShellCollection and its first element exist before modification.
            //if (_tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells != null &&
            //    _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells.Count > 0)
            //{
            //    // Changes the length of the cylindrical shell based on the total calculated length.
            //    _tankSiteAssembly._assemblyOfCylindricalShells.ActivateDocument();
            //    _tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[0].ChangeLength(cylindricalShellLength);
            //}
            //else
            //{
            //    MessageBox.Show("Error adjusting cylindrical shell length.");
            //}

            //// Adjust nozzles
            //AdjustNozzleCompartments(COMPARTMENT_DOC_PATH, compartmentsLengths);
        }

        /// <summary>
        /// Depending on which radio button (default or select project folder) is selected:
        /// Opens folder browser window for user to select a different project folder. Or
        /// Sets the primary default project folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectFolderRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            // Let the user to select a different project
            if (SelectFolderRadioButton.Checked)
            {
                folderBrowserDialog1.ShowDialog();
                _folderForAllProjects = folderBrowserDialog1.SelectedPath;

                SetProjectsFolder();
            }
            // Set project folder to the default folder
            else
                GetProjectsFolder();
        }

        ///// <summary>
        ///// Changes visibility of manhole controls depending on the state of default manhole configuration check box
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void DefaultManholeConfigCheckBox_CheckedChanged(object sender, EventArgs e)
        //{
        //    // If default manhole configuration check box is unchecked,
        //    // hide manholes controls
        //    if (!defaultManholeConfigCheckBox.Checked)
        //    {
        //        manholesPanel.Visible = false;
        //        manholeSettingPanel.Visible = false;
        //    }

        //    // If default manhole configuration check box is checked,
        //    // display manholes controls
        //    else
        //    {
        //        manholesPanel.Visible = true;
        //        manholeSettingPanel.Visible = true;
        //    }
        //}

        private void TypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            int selectedIndex = comboBox.SelectedIndex;

            if (selectedIndex == 1)
                DoubleSkinParametersPanel.Visible = true;

            else
                DoubleSkinParametersPanel.Visible = false;
        }
    }
}

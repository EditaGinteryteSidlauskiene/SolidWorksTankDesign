using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Helpers;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace SolidWorksTankDesign.MVP.Presenters
{
    public class InitialConfiguarationPresenter
    {
        private const string DISHED_END_MAIN_DOC_NAME = "Main dished end.SLDPRT";
        private const string INNER_DISHED_END_DOC_NAME = "Inner dished end.SLDPRT";
        private const string CYLINDRICAL_SHELL_DOC_NAME = "Cylindrical shell.SLDPRT";
        private const string EMPTY_TANKSITE_ASSEMBLY_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Tank Site Assembly.SLDASM";
        private const string DISHED_END_VOLUME_INCLUDE_PROP_NAME = "volume include";
        private const string DISHED_END_VOLUME_EXCLUDE_PROP_NAME = "volume exclude";
        private const string CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME = "volume per meter";
        private const double INITIAL_COMPARTMENT_LENGTH = 2000;

        private readonly IInitialConfigurationView _initialConfigurationView;

        private string _settingsFilePath;
        private string _folderForAllProjects;
        private string _projectFolder;
        private  string _mainFolderPath;
        private string _compartmentDocPath;

        private double _mainDishedEndIncludeVolume;
        private double _innerDishedEndExcludeVolume;
        private double _innerDishedEndIncludeVolume;
        private double _cylindricalShellVolumePerMeter;
        private string _instalation;
        private double _diameter;

        public InitialConfiguarationPresenter(
            IInitialConfigurationView initialConfigurationView, 
            string settingsFilePath, 
            string mainFolderPath,
            string compartmentDocPath)
        {
            _initialConfigurationView = initialConfigurationView;
            _settingsFilePath = settingsFilePath;
            _mainFolderPath = mainFolderPath;
            _compartmentDocPath = compartmentDocPath;

            _initialConfigurationView.PopulateStandardsComboBox(GetConstructionStandardDescriptions());
            _initialConfigurationView.PopulateTankTypeComboBox(GetTankTypeDescriptions());
            _initialConfigurationView.PopulateClassComboBox((Class[])Enum.GetValues(typeof(Class)));
            _initialConfigurationView.PopulateExternalTreatmentsComboBox(GetExternalTreatmentsList());

            _initialConfigurationView.SaveInitialConfiguration += _initialConfigurationView_SaveInitialConfiguration;

            _folderForAllProjects = DocumentManager.GetProjectsFolder(settingsFilePath);
        }

        private void GetInstalation()
        {
            if (_initialConfigurationView.Standard == ConstructionStandard.EN_12285_1)
                _instalation = "u";
            else
                _instalation = "a";
        }

        private void GetDiameter()
        {
            double.TryParse(_initialConfigurationView.Diameter.Replace(',', '.').Replace(" ", ""), out _diameter);
        }

        private void _initialConfigurationView_SaveInitialConfiguration(object sender, EventArgs e)
        {
            if (!AreUserInputsValid()) return;

            ViewsHelper.DisableControls((Control)_initialConfigurationView);

            GetInstalation();
            GetDiameter();

            AddInitialTank();

            AssignTankInitialConfigurations();

            ICompartmentWindowView compartmentView = new CompartmentWindowView(_projectFolder, _mainFolderPath, _compartmentDocPath);

            _initialConfigurationView.ShowCompartmentsWindow((Control)compartmentView);
        }

        private Dictionary<string, ConstructionStandard> GetConstructionStandardDescriptions()
        {
            // 1. Create a dictionary to map descriptions to enum values
            var standardMap = new Dictionary<string, ConstructionStandard>();

            // 2. Populate the dictionary with descriptions and enum values
            foreach (var name in Enum.GetNames(typeof(ConstructionStandard)))
            {
                ConstructionStandard standard = (ConstructionStandard)Enum.Parse(typeof(ConstructionStandard), name);
                string description = EnumManager.GetEnumDescription(standard);
                standardMap[description] = standard;
            }

            return standardMap;
        }

        private Dictionary<string, TankType> GetTankTypeDescriptions()
        {
            // 1. Create a dictionary to map descriptions to enum values
            var tankTypeMap = new Dictionary<string, TankType>();

            // 2. Populate the dictionary with descriptions and enum values
            foreach (var name in Enum.GetNames(typeof(TankType)))
            {
                TankType tankType = (TankType)Enum.Parse(typeof(TankType), name);
                string description = EnumManager.GetEnumDescription(tankType);
                tankTypeMap[description] = tankType;
            }

            return tankTypeMap;
        }

        private List<Treatment> GetExternalTreatmentsList()
        {
            return SettingsDataManager.GetTreatments()
                .Where(t => t.Type == TreatmentType.ExternalUnderGround || t.Type == TreatmentType.ExternalAboveGround).ToList();
        }

        private bool IsSerialNumberInputValid()
        {
            bool result = true;
            string serialNumberText = _initialConfigurationView.SerialNumber;

            // Check if serial number is entered
            if (serialNumberText.Trim() == string.Empty)
            {
                _initialConfigurationView.LocateRedDotNextToSerialNumberTextBox();

                result = false;
            }
            // Check if there are no forbidden characters in serial no. These characters cannot be used in folder's name
            else if (serialNumberText.Contains('<') || serialNumberText.Contains('>') || serialNumberText.Contains(':') ||
                serialNumberText.Contains('"') || serialNumberText.Contains('/') || serialNumberText.Contains('\\') ||
                serialNumberText.Contains('|') || serialNumberText.Contains('?') || serialNumberText.Contains('*'))
            {
                MessageBox.Show("Characters that are not allowed in Serial no.:\n <   >   :   \"   /   \\   |   ?   *");
                _initialConfigurationView.SerialNumber = string.Empty;

                result = false;
            }
            // Check if serial no. is not a reserved name for folders
            else if (serialNumberText == "CON" || serialNumberText == "PRN" || serialNumberText == "AUX" ||
                serialNumberText == "NUL" || serialNumberText == "COM1" || serialNumberText == "COM2" ||
                serialNumberText == "COM3" || serialNumberText == "COM4" || serialNumberText == "COM5" ||
                serialNumberText == "COM6" || serialNumberText == "COM7" || serialNumberText == "COM8" ||
                serialNumberText == "COM9" || serialNumberText == "LPT1" || serialNumberText == "LPT2" ||
                serialNumberText == "LPT3" || serialNumberText == "LPT4" || serialNumberText == "LPT5" ||
                serialNumberText == "LPT6" || serialNumberText == "LPT7" || serialNumberText == "LPT8" ||
                serialNumberText == "LPT9")
            {
                MessageBox.Show("Serial no. cannot be:\n \"CON\", \"PRN\", \"AUX\", \"NUL\", \"COM1\", \"COM2\", \"COM3\", \"COM4\", \"COM5\"" +
                    ", \"COM6\", \"COM7\", \"COM8\", \"COM9\", \"LPT1\", \"LPT2\", \"LPT3\", \"LPT4\", \"LPT5\", \"LPT6\", \"LPT7\", \"LPT8\"" +
                    ", \"LPT9\"");
                _initialConfigurationView.SerialNumber = string.Empty;

                result = false;
            }
            // Check if serial no does not start or end with ".". This is not allowed in folder's name.
            else if (serialNumberText.StartsWith(".") || serialNumberText.EndsWith("."))
            {
                MessageBox.Show("Serial no. cannot start or end with \".\"");
                _initialConfigurationView.SerialNumber = string.Empty;

                result = false;
            }

            return result;
        }

        private bool AreUserInputsValid()
        {
            bool result = true;

            if (!IsSerialNumberInputValid())
            {
                result = false;
            }
                

            if (!double.TryParse(_initialConfigurationView.Diameter.Replace(" ", "").Replace(',', '.'), out _))
            {
                _initialConfigurationView.LocateRedDotNextToDiameterTextBox();
                result = false;
            }
                

            if (!double.TryParse(_initialConfigurationView.MinOperatingTemp.Replace(" ", "").Replace(',', '.'), out _))
            {
                _initialConfigurationView.LocateRedDotNextToMinOperatingTempTextBox();
                result = false;
            }
               

            if (!double.TryParse(_initialConfigurationView.MaxOperatingTemp.Replace(" ", "").Replace(',', '.'), out _))
            {
                _initialConfigurationView.LocateRedDotNextToMaxOperatingTempTextBox();
                result = false;
            }

            if (!double.TryParse(_initialConfigurationView.MaxOperatingPressure.Replace(" ", "").Replace(',', '.'), out _))
            {
                _initialConfigurationView.LocateRedDotNextToMaxOperatingPressureTextBox();
                result = false;
            }


            if (!double.TryParse(_initialConfigurationView.ShellLeakTestPressure.Replace(" ", "").Replace(',', '.'), out _))
            {
                _initialConfigurationView.LocateRedDotNextToShellLeakTestPressureTextBox();
                result = false;
            }

            return result;
        }

        

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
                $"{_instalation.ToUpper()}T",
                $"{_instalation.ToUpper()}T {_diameter} Class {_initialConfigurationView.Class}");

            // Get the path of that folder. It is stored in the main folder
            string folderPath = Path.Combine(_mainFolderPath, folderName);

            // If the folder exists
            if (Directory.Exists(folderPath))
            {
                // Return the full document path
                return Path.Combine(folderPath, documentName);
            }

            // Else, message that the folder was not found
            else
            {
                MessageBox.Show($"Folder {folderName} was not found in {_mainFolderPath} path.");
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
                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.CompleteMainDishedEnds(_projectFolder, dishedEndDocPath);
            }

            else
                return;
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
                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells.Add(new CylindricalShell());
                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells._cylindricalShellDocPath = cylindricalShellDocPath;
                SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells.CylindricalShells[0].CompleteFirstCylindricalShel(_projectFolder, length);
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
            CompartmentsManager compartmentsManager = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager;
            // Activate shell document
            ModelDoc2 shellDoc = compartmentsManager.ActivateDocument();

            Feature leftDishedEndPositionPlane;

            // Get assembly of dished ends document
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetDishedEndsAssemblyComponent().GetModelDoc2();
            // Get left dished end's position plane
            using (var dishedEndsDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, assemblyOfDishedEndsDoc))
            {
                leftDishedEndPositionPlane = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.LeftDishedEnd.GetPositionPlane();
            }

            // Add compartment
            compartmentsManager.Compartments.Add(new Compartment(
                _projectFolder,
                _compartmentDocPath,
                SWFeatureManager.GetMajorPlane(shellDoc, MajorPlane.Front),
                compartmentsManager.GetCenterAxis(),
                leftDishedEndPositionPlane,
                length));

            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0]._compartmentSettings.ID = Guid.NewGuid();

            DocumentManager.UpdateAndSaveDocuments();
        }

        private void AddTankParts()
        {
            if (SolidWorksDocumentProvider._tankSiteAssembly == null)
            {
                MessageBox.Show("Error creating tank site assembly object.");
                return;
            }

            // Get custom properties values
            _mainDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(DISHED_END_MAIN_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _innerDishedEndExcludeVolume = -SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_EXCLUDE_PROP_NAME);
            _innerDishedEndIncludeVolume = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(INNER_DISHED_END_DOC_NAME), DISHED_END_VOLUME_INCLUDE_PROP_NAME);
            _cylindricalShellVolumePerMeter = SWFeatureManager.GetCustomPropertyValue(GetDocumentPath(CYLINDRICAL_SHELL_DOC_NAME), CYLINDRICAL_SHELL_VOLUME_PER_METER_PROP_NAME);

            AddMainDishedEnds();

            AddFirstCylindricalShell(INITIAL_COMPARTMENT_LENGTH / 1000);
            AddFirstCompartment(INITIAL_COMPARTMENT_LENGTH / 1000);

            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.ActivateDocument();
            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(INITIAL_COMPARTMENT_LENGTH / 1000);
            SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds.CloseDocument();
        }

        private void CreateFilesToDeleteList()
        {
            SolidWorksDocumentProvider._filesToDelete = new List<string>();
        }

        private void AddInitialCompartmentConfiguration(double initialCompartmentVolume)
        {
            SolidWorksDocumentProvider._tankProperties = new TankProperties();

            SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Add(new CompartmentConfiguration
            {
                ID = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0]._compartmentSettings.ID,
                Name = "Compartment A",
                Volume = initialCompartmentVolume,
                Length = INITIAL_COMPARTMENT_LENGTH,
                LeftDishedEndAlignment = DishedEndAlignment.Left,
                LeftEndConnection = LeftEndConnection.None
            });

            TankSiteDataManager.UpdateTankProperties();

            CompartmentsMappingHelper.AddOrUpdateMapping(
                SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0],
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[0]);
        }

        private void AddInitialTank()
        {
            SldWorks solidWorksApp = SolidWorksDocumentProvider._solidWorksApplication;

            try
            {
                // Open empty tank site asssembly doc
                DocumentSpecification docSpecification =
                    (DocumentSpecification)solidWorksApp.GetOpenDocSpec(EMPTY_TANKSITE_ASSEMBLY_DOC_PATH);
                ModelDoc2 emptyTankSiteAssemblyDoc = solidWorksApp.OpenDoc7(docSpecification);

                // Get serial number, class, and number of compartments
                string serialNumber = _initialConfigurationView.SerialNumber;
                Class selectedClass = _initialConfigurationView.Class;

                _projectFolder = DocumentManager.CopyDocuments(solidWorksApp, emptyTankSiteAssemblyDoc, _folderForAllProjects, serialNumber);

                // Initialize and store tank site configurations
                SolidWorksDocumentProvider._tankSiteAssembly = new TankSiteAssembly();
                SolidWorksDocumentProvider._tankSiteAssembly.InitializeAndStoreTankSiteConfiguration();

                // Configure tank shell and compartments according tank configuration the user entered
                AddTankParts();

                //Recalculate volume
                double cylindricalShellVolume = _cylindricalShellVolumePerMeter * 2;
                double initialCompartmentVolume = cylindricalShellVolume + (_mainDishedEndIncludeVolume * 2);

                AddInitialCompartmentConfiguration(initialCompartmentVolume);

                CreateFilesToDeleteList();

                DocumentManager.UpdateAndSaveDocuments();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void AssignTankInitialConfigurations()
        {
            TankProperties tankProperties = SolidWorksDocumentProvider._tankProperties;

            tankProperties.SerialNumber = _initialConfigurationView.SerialNumber;

            tankProperties.ConstructionStandard = _initialConfigurationView.Standard;

            tankProperties.Type = _initialConfigurationView.TankType;

            tankProperties.Class = _initialConfigurationView.Class;

            double.TryParse(_initialConfigurationView.Diameter.Replace(" ", "").Replace(',', '.'), out double diameter);
            tankProperties.NominalDiameter = diameter;

            double.TryParse(_initialConfigurationView.MinOperatingTemp.Replace(" ", "").Replace(',', '.'), out double minOperatingTemp);
            tankProperties.MinOperatingTemperature = minOperatingTemp;

            double.TryParse(_initialConfigurationView.MaxOperatingTemp.Replace(" ", "").Replace(',', '.'), out double maxOperatingTemp);
            tankProperties.MaxOperatingTemperature = maxOperatingTemp;

            double.TryParse(_initialConfigurationView.MaxOperatingPressure.Replace(" ", "").Replace(',', '.'), out double maxOperatingPressure);
            tankProperties.MaxOperatingPressure = maxOperatingPressure;

            double.TryParse(_initialConfigurationView.ShellLeakTestPressure.Replace(" ", "").Replace(',', '.'), out double shellLeakTestPressure);
            tankProperties.ShellLeakTestPressure = shellLeakTestPressure;

            tankProperties.ExternalSurfaceTreatment = _initialConfigurationView.ExternalSurfaceTreatment;

            TankSiteDataManager.UpdateTankProperties();
        }
    }
}

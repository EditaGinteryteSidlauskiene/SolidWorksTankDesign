using Newtonsoft.Json;
using SolidWorksTankDesign;
using System.Windows.Forms;
using System;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.Windows;
using System.IO;
using SolidWorksTankDesign.Helpers;
using System.Threading.Tasks;

namespace MVP
{
    public class TaskpaneHostUIPresenter
    {
        private const string SETTINGS_FILE_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Settings.txt";
        private const string MAIN_FOLDER_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite";
        private const string COMPARTMENT_DOC_PATH = "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Compartment.SLDASM";

        private readonly ITaskpaneHostUI _taskpaneHostUIView;

        public TaskpaneHostUIPresenter(ITaskpaneHostUI taskpaneHostUIView)
        {
            _taskpaneHostUIView = taskpaneHostUIView;

            _taskpaneHostUIView.UpdateSettings += _taskpaneHostUIView_UpdateSettings;
            _taskpaneHostUIView.CreateTank += _taskpaneHostUIView_CreateTank;
            _taskpaneHostUIView.RecognizeTankAssemlby += _taskpaneHostUIView_RecognizeTankAssemlby;
        }

        private async void _taskpaneHostUIView_RecognizeTankAssemlby(object sender, EventArgs e)
        {
            SolidWorksDocumentProvider._tankSiteAssembly = TankSiteDataManager.LoadTankSiteAssemblyFromAttribute();
            SolidWorksDocumentProvider._tankProperties = TankSiteDataManager.LoadTankSiteAssemblyPropertiesFromAttribute();
            SolidWorksDocumentProvider._filesToDelete = TankSiteDataManager.LoadFilesToDeleteFromAttribute();

            if (SolidWorksDocumentProvider._filesToDelete.Count > 0) DocumentManager.DeleteFiles();

            for (int i = 0; i < SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments.Count; i++)
                CompartmentsMappingHelper.AddOrUpdateMapping(
                SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[i],
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i]);

            ICompartmentWindowView compartmentView = new CompartmentWindowView(
                Path.GetDirectoryName(SolidWorksDocumentProvider.GetActiveDoc().GetPathName()), 
                MAIN_FOLDER_PATH, 
                COMPARTMENT_DOC_PATH);

            _taskpaneHostUIView.ShowCompartmentsWindow((Control)compartmentView);
        }

        private void _taskpaneHostUIView_CreateTank(object sender, EventArgs e)
        {
            IInitialConfigurationView initialConfigurationView = new InitialConfigurationView(
                SETTINGS_FILE_PATH, 
                MAIN_FOLDER_PATH,
                COMPARTMENT_DOC_PATH);

            _taskpaneHostUIView.ShowInitialConfigurationWindow((Control)initialConfigurationView);
        }

        private void _taskpaneHostUIView_UpdateSettings(object sender, EventArgs e)
        {

        }

       
    
    }
}

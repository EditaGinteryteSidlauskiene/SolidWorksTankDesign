using SolidWorksTankDesign;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.MVP.Views.Controls;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Globalization;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Presenters
{
    public class NozzleWindowPresenter
    {
        INozzleWindowView _nozzleWindowView;
        INozzleModel _nozzleModel;
        ICompartmentConfigurationModel _compartmentConfigurationModel;

        public NozzleWindowPresenter(string projectFolderPath, INozzleWindowView nozzleWindowView, INozzleModel nozzleModel, ICompartmentConfigurationModel compartmentConfigurationModel)
        {
            _nozzleWindowView = nozzleWindowView;
            _nozzleModel = nozzleModel;
            _compartmentConfigurationModel = compartmentConfigurationModel;

            _nozzleWindowView.NewNozzleButtonClicked += OnNewNozzleButtonClicked;
            _nozzleWindowView.ForwardButtonPressed += OnForwardButtonPressed;
            _nozzleWindowView.BackButtonPressed += OnBackButtonPressed;
        }

        private void OnBackButtonPressed(object sender, EventArgs e)
        {
            // Navigate back to compartment window
            // This will be implemented based on your navigation pattern
        }

        private void OnForwardButtonPressed(object sender, EventArgs e)
        {
            // Update tank properties with all configuration changes (including nozzles)
            TankSiteDataManager.UpdateTankProperties();

            // Apply nozzle changes to the actual Nozzle objects in SolidWorks
            _nozzleModel.ApplyNozzleChanges();

            // TODO: Navigate to next window or complete the configuration
        }

        private void OnNewNozzleButtonClicked(object sender, EventArgs e)
        {
            Panel compartmentPanel = sender as Panel;
            CompartmentConfiguration compartmentConfiguration = (CompartmentConfiguration)compartmentPanel.Tag;

            NozzleConfiguration newNozzleConfig = new NozzleConfiguration();
            compartmentConfiguration.NozzleConfigurations.Add(newNozzleConfig);

            _nozzleWindowView.AddNozzlePanel(newNozzleConfig, compartmentPanel);
        }

        public void UpdateReferencePoint(NozzleConfiguration nozzleConfiguration, NozzleConfigurationControl.DistanceChangedEventArgs e)
        {
            if (nozzleConfiguration == null || e == null) return;

            if (e.TopReferenceType.HasValue)
            {
                nozzleConfiguration.TopReferenceType = e.TopReferenceType.Value;
            }
            else if (e.BottomReferencePoint.HasValue)
            {
                nozzleConfiguration.BottomReferencePoint = e.BottomReferencePoint.Value;
            }
        }
    }
}

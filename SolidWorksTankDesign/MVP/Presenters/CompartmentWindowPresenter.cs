using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Helpers;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Presenters
{
    public class CompartmentWindowPresenter
    {
        ICompartmentWindowView _compartmentWindowView;
        ICompartmentConfigurationModel _compartmentConfigurationModel;

        public CompartmentWindowPresenter(ICompartmentWindowView compartmentWindowView, ICompartmentConfigurationModel compartmentConfigurationModel) 
        {
            _compartmentWindowView = compartmentWindowView;
            _compartmentConfigurationModel = compartmentConfigurationModel;

            _compartmentWindowView.NewCompartmentButtonClicked += OnNewCompartmentButtonClicked;
            _compartmentWindowView.PaintingSystemComboBoxSelectionChanged += OnPaintingSystemComboBoxSelectionChanged;
            _compartmentWindowView.ForwardButtonPressed += OnForwardButtonPressed;
            _compartmentWindowView.BackButtonPressed += OnBackButtonPressed;
            ProvideInternalTreatmentBindingSource();
            
            _compartmentWindowView.GetPaintingAmountUnitsDictionary(GetPaintingAmountUnitsDescriptions());

            LoadCompartments();
        }

        private void OnBackButtonPressed(object sender, EventArgs e)
        {
            _compartmentConfigurationModel.BackButtonClick();
        }

        public void ShowPaintingSystemsManagerView(CompartmentWindowPresenter compartmentWindowPresenter)
        {
            IPaintingSystemsManagerView paintingSystemsManagerView = new PaintingSystemsManagerView(
                _compartmentWindowView,
                compartmentWindowPresenter,
                TreatmentType.Internal);
            _compartmentWindowView.ShowPaintingSystemsManagerView((Control)paintingSystemsManagerView);
        }

        public void RemoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration)
        {
            _compartmentConfigurationModel.RemoveCompartmentConfiguration(compartmentConfiguration);
        }

        public void CalculateVolume(CompartmentConfiguration compartmentConfiguration)
        {
            _compartmentConfigurationModel.CalculateVolume(compartmentConfiguration);
        }

        public void CalculateLength(CompartmentConfiguration compartmentConfiguration)
        {
            _compartmentConfigurationModel.CalculateLength(compartmentConfiguration);
        }

        private void OnForwardButtonPressed(object sender, EventArgs e)
        {
            TankSiteDataManager.UpdateTankProperties();

            ViewsHelper.DisableControls((Control)_compartmentWindowView);

            _compartmentConfigurationModel.ModifyTank();

            ViewsHelper.EnableControls((Control)_compartmentWindowView);

            INozzleWindowView nozzleView = new NozzleWindowView((CompartmentConfigurationModel)_compartmentConfigurationModel);

            _compartmentWindowView.ShowNozzleWindow((Control)nozzleView);
        }

        public void RenameCompartments()
        {
            _compartmentConfigurationModel.AdjustCompartmentsNames();
        }

        public void ChangeLeftDishedEndConnection(CompartmentConfiguration compartmentConfiguration, LeftEndConnection newConnection)
        {
            _compartmentConfigurationModel.UpdateLeftEndConnection(compartmentConfiguration, newConnection);
        }

        public void ChangeLeftDishedEndAlignment(CompartmentConfiguration compartmentConfiguration, DishedEndAlignment newAlignment)
        {
            _compartmentConfigurationModel.ChangeLeftDishedEndAlignment(compartmentConfiguration, newAlignment);
        }

        public void MoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration, int newIndex)
        {
            _compartmentConfigurationModel.MoveCompartmentConfiguration(compartmentConfiguration, newIndex);
        }

        private void ProvideInternalTreatmentBindingSource()
        {
            Container components = new Container();
            BindingSource internalTreatmentBindingSource = new BindingSource(components);
            internalTreatmentBindingSource.DataSource = _compartmentConfigurationModel.InternalTreatments;

            _compartmentWindowView.GetInternalTreatmentsBindingSource(internalTreatmentBindingSource);

            _compartmentWindowView.UpdatePaintingSystemsComboBoxes();
        }

        private void OnPaintingSystemComboBoxSelectionChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            CompartmentConfiguration compartmentConfiguration = comboBox.Tag as CompartmentConfiguration;

            Treatment selectedInternalTreatment = comboBox.SelectedItem as Treatment;
            compartmentConfiguration.InternalSurfaceTreatment = selectedInternalTreatment;
        }

        private void LoadCompartments()
        {
            _compartmentWindowView.ClearCompartmentPanels(); // Clear any existing panels

            for(int i = 0; i < _compartmentConfigurationModel.CompartmentConfigurations.Count; i++)
            {
                _compartmentWindowView.AddCompartmentPanel(_compartmentConfigurationModel.CompartmentConfigurations[i], i);
            }

        }

        private void OnNewCompartmentButtonClicked(object sender, EventArgs e)
        {
            CompartmentConfiguration newCompartmentConfig = _compartmentConfigurationModel.CreateCompartmentConfiguration();

            _compartmentWindowView.AddCompartmentPanel(newCompartmentConfig, _compartmentConfigurationModel.CompartmentConfigurations.Count - 1);
        }

        public void UpdateInternalTreatmentBindingSource()
        {
            _compartmentConfigurationModel.GetInternalTreatments();

            ProvideInternalTreatmentBindingSource();
        }

        public Dictionary<string, PaintingAmountUnit> GetPaintingAmountUnitsDescriptions()
        {
            // 1. Create a dictionary to map descriptions to enum values
            var standardMap = new Dictionary<string, PaintingAmountUnit>();

            // 2. Populate the dictionary with descriptions and enum values
            foreach (var name in Enum.GetNames(typeof(PaintingAmountUnit)))
            {
                PaintingAmountUnit unit = (PaintingAmountUnit)Enum.Parse(typeof(PaintingAmountUnit), name);
                string description = EnumManager.GetEnumDescription(unit);
                standardMap[description] = unit;
            }

            return standardMap;
        }
    }
}

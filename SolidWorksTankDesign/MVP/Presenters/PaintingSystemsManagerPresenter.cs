using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Presenters
{
    internal class PaintingSystemsManagerPresenter
    {
        IPaintingSystemsModel _paintingSystemsModel;
        IPaintingSystemsManagerView _paintingSystemsManagerView;
        CompartmentWindowPresenter _compartmentWindowPresenter;

        private TreatmentType _treatmentType;

        public ObservableCollection<Treatment> Treatments { get; private set; } =  new ObservableCollection<Treatment>();

        public PaintingSystemsManagerPresenter (
            CompartmentWindowPresenter compartmentWindowPresenter,
            IPaintingSystemsManagerView paintingSystemsManagerView, 
            IPaintingSystemsModel paintingSystemsModel,
            TreatmentType treatmentType)
        {
            _paintingSystemsManagerView = paintingSystemsManagerView;
            _paintingSystemsModel = paintingSystemsModel;
            _compartmentWindowPresenter = compartmentWindowPresenter;

            _treatmentType = treatmentType;
            LoadTreatments(_treatmentType);

            _paintingSystemsManagerView.TreatmentType = _treatmentType;
            _paintingSystemsManagerView.TreatmentListDataSource = Treatments;
            _paintingSystemsManagerView.TreatmentSelected += OnTreatmentSelected;
            _paintingSystemsManagerView.DeleteLayer += OnDeleteLayer;
            _paintingSystemsManagerView.DeleteTreatment += OnDeleteTreatment;
            _paintingSystemsManagerView.GoBack += OnGoBack;
            _paintingSystemsManagerView.SaveAllChanges += OnSaveAllChanges;
            _paintingSystemsManagerView.CancelAllChanges += OnCancelAllChanges;

            _paintingSystemsModel.BeginEdit();
        }

        private void OnSaveAllChanges(object sender, EventArgs e)
        {
            Treatment treatmentToValidate;

            if (_paintingSystemsManagerView.SelectedTreatmentIndex >= 0)
                treatmentToValidate = Treatments[_paintingSystemsManagerView.SelectedTreatmentIndex];

            else
            {
                treatmentToValidate = _paintingSystemsManagerView.TempTreatment;
            }

            treatmentToValidate.Description = _paintingSystemsManagerView.TreatmentDescription;
            treatmentToValidate.Cleaning = _paintingSystemsManagerView.TreatmentCleaning;
            treatmentToValidate.Comments = _paintingSystemsManagerView.TreatmentComments;

            // Validate all treatments before saving
            var errors = ValidateTreatment(treatmentToValidate);
            if (errors.Any())
            {
                _paintingSystemsModel.CancelEdit();
                _paintingSystemsManagerView.ShowErrors(errors);

                LoadTreatments(_treatmentType);
            }
            else
            {
                // If it's a brand-new treatment not in the list
                if (!Treatments.Contains(treatmentToValidate))
                {
                    Treatments.Insert(Treatments.Count - 1, treatmentToValidate);
                    _paintingSystemsModel.SaveNewTreatment(treatmentToValidate);

                    _paintingSystemsManagerView.RefreshBidnings(Treatments);
                    _paintingSystemsManagerView.SelectAndDisplayNewlyAddedTreatment();

                    _paintingSystemsModel.BeginEdit();
                    _paintingSystemsModel.EndEdit();
                }
                else
                {
                    _paintingSystemsModel.BeginEdit();
                    _paintingSystemsModel.EndEdit();

                    _paintingSystemsManagerView.RefreshBidnings(Treatments);
                }
            }
        }

        private void OnCancelAllChanges(object sender, EventArgs e)
        {
            _paintingSystemsModel.CancelEdit();
            _paintingSystemsManagerView.ShowMessage("All changes have been canceled.");
            // Refresh the UI to reflect the reverted state
        }

        private void OnGoBack(object sender, EventArgs e)
        {
            _compartmentWindowPresenter.UpdateInternalTreatmentBindingSource();
        }

        private void OnDeleteTreatment(object sender, EventArgs e)
        {
            Treatment treatment = Treatments[_paintingSystemsManagerView.SelectedTreatmentIndex];

            Treatments.Remove(treatment);
            _paintingSystemsModel.DeleteTreatment(treatment);

            _paintingSystemsManagerView.RefreshBidnings(Treatments);
        }

        private void OnDeleteLayer(object sender, EventArgs e)
        {
            Treatment treatment = Treatments[_paintingSystemsManagerView.SelectedTreatmentIndex];

            _paintingSystemsModel.DeleteLayer(
                treatment, _paintingSystemsManagerView.GetSelectedCoatingLayerIndices());

            _paintingSystemsManagerView.RefreshCoatingLayersDataGridView(treatment);
        }

        private void OnTreatmentSelected(object sender, EventArgs e)
        {
            int selectedIndex = _paintingSystemsManagerView.SelectedTreatmentIndex;
            if (selectedIndex >= 0 && selectedIndex < Treatments.Count)
            {
                Treatment treatment = Treatments[selectedIndex];
                _paintingSystemsManagerView.TreatmentDescription = treatment.Description;
                _paintingSystemsManagerView.TreatmentCleaning = treatment.Cleaning;
                _paintingSystemsManagerView.TreatmentComments = treatment.Comments;

                _paintingSystemsManagerView.DisplaySelectedTreatmentDetails(treatment);
            }
            else
            {
                // Clear TextBox values if no selection
                _paintingSystemsManagerView.TreatmentDescription = string.Empty;
                _paintingSystemsManagerView.TreatmentCleaning = string.Empty;
                _paintingSystemsManagerView.TreatmentComments = string.Empty;
            }
        }

        private void LoadTreatments(TreatmentType treatmentType)
        {
            Treatments.Clear();

            List<Treatment> treatmentsOfType = _paintingSystemsModel.Treatments.ToList()
                    .Where(t => t.Type == treatmentType).ToList();

            foreach (Treatment treatment in treatmentsOfType)
            {
                Treatments.Add(treatment);
            }

            Treatments.Add(new Treatment { Description = "New Treatment" });
        }

        /// <summary>
        /// Validates all treatments and their coating layers.
        /// </summary>
        /// <returns>Dictionary of errors with property names as keys and error messages as values.</returns>
        public Dictionary<string, string> ValidateTreatment(Treatment treatment)
        {
            var errors = new Dictionary<string, string>();

            if (string.IsNullOrWhiteSpace(treatment.Description))
                errors.Add("Description", "Treatment description cannot be empty.");

            // 2. Ensure the description is unique among all treatments except itself
            bool isDuplicateDescription = _paintingSystemsModel.Treatments
                .Any(t => t != treatment && t.Description.Equals(treatment.Description, StringComparison.OrdinalIgnoreCase));

            if (isDuplicateDescription)
                errors.Add("Description", $"Treatment with {treatment.Description} description already exists.");

            if (string.IsNullOrWhiteSpace(treatment.Cleaning))
                errors.Add("Cleaning", "Treatment cleaning cannot be empty.");

            if(treatment.CoatingLayers.Count == 0)
                errors.Add("Coating Layers", "Coating layers cannot be empty.");

            foreach (var layer in treatment.CoatingLayers)
            {
                if (string.IsNullOrWhiteSpace(layer.ProductName))
                    errors.Add("ProductName", "Product Name in Coating Layer cannot be empty.");

                if (string.IsNullOrWhiteSpace(layer.Units))
                    errors.Add("Units", "Units in Coating Layer cannot be empty.");
            }

            return errors;
        }

        public bool ValidateRow(DataGridViewRow dataGridViewRow)
        {
            foreach (DataGridViewCell cell in dataGridViewRow.Cells)
            {
                var columnName = dataGridViewRow.DataGridView.Columns[cell.ColumnIndex].DataPropertyName;
                var value = cell.Value?.ToString();

                // Perform validation for specific columns
                if (columnName == "DryFilmThickness" || columnName == "WetFilmThickness" || columnName == "FactualConsumption")
                {
                    if (!double.TryParse(value, out double result) || result < 0)
                    {
                        return false; // Invalid numeric input
                    }
                }
                else if (columnName == "ProductName" || columnName == "Units")
                {
                    if (string.IsNullOrWhiteSpace(value) || value == "Product Name" || value == "Units")
                    {
                        return false; // Empty string
                    }
                }
            }

            return true; // All cells are valid
        }
    }
}

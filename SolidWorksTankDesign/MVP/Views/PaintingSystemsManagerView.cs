using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Presenters;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Control = System.Windows.Forms.Control;
using UserControl = System.Windows.Forms.UserControl;

namespace SolidWorksTankDesign.MVP.Views
{
    public partial class PaintingSystemsManagerView : UserControl, IPaintingSystemsManagerView
    {
        PaintingSystemsManagerPresenter _paintingSystemsManagerPresenter;
        ICompartmentWindowView _compartmentWindowView;

        private Treatment _tempTreatment;

        private BindingList<CoatingLayer> _coatingLayers { get; set; } = new BindingList<CoatingLayer>();

        private readonly ErrorProvider errorProvider = new ErrorProvider();

        public PaintingSystemsManagerView(
            ICompartmentWindowView compartmentWindowView,
             CompartmentWindowPresenter compartmentWindowPresenter,
             TreatmentType treatmentType)
        {
            InitializeComponent();
            _compartmentWindowView = compartmentWindowView;

            IPaintingSystemsModel paintingSystemsModel = new PaintingSystemsModel();
            _paintingSystemsManagerPresenter = new PaintingSystemsManagerPresenter(compartmentWindowPresenter, this, paintingSystemsModel, treatmentType);

            TreatmentsListBox.SelectedIndex = -1;
        }

        public TreatmentType TreatmentType { get; set; }

        public Treatment TempTreatment
        {
            get { return _tempTreatment; }
            set { _tempTreatment = value; }
        }

        public string TreatmentDescription
        {
            get => TreatmentDescriptionTextBox.Text;
            set => TreatmentDescriptionTextBox.Text = value;
        }

        public string TreatmentCleaning
        {
            get => TreatmentCleaningTextBox.Text;
            set => TreatmentCleaningTextBox.Text = value;
        }

        public string TreatmentComments
        {
            get => TreatmentCommentsTextBox.Text;
            set => TreatmentCommentsTextBox.Text = value;
        }

        private void ResizeDataGridViewToFitRows(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count == 0)
            {
                // Handle case where no rows are present
                dataGridView.Height = dataGridView.ColumnHeadersHeight;
                return;
            }

            // Calculate the total height of visible rows
            int totalRowHeight = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.Visible)
                {
                    totalRowHeight += row.Height;
                }
            }

            // Add the height of the "add new" row if applicable
            if (dataGridView.AllowUserToAddRows)
            {
                totalRowHeight += dataGridView.Rows[dataGridView.Rows.Count - 1].Height;
            }

            // Add the height of the column headers
            int totalHeight = totalRowHeight + dataGridView.ColumnHeadersHeight;

            // Add padding for gridlines or borders (optional)
            totalHeight += 2; // Small adjustment for borders

            // Limit the total height to a maximum value if necessary
            const int maxHeight = 150; // Example maximum height
            if (totalHeight > maxHeight)
            {
                totalHeight = maxHeight;
            }

            // Set the new height of the DataGridView
            dataGridView.Height = totalHeight;

            // Calculate the total width of visible columns
            int totalWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible);

            // Add the width of row headers (if visible)
            totalWidth += dataGridView.RowHeadersWidth;

            // Set the new width of the DataGridView
            dataGridView.Width = totalWidth + 2; // Small adjustment for borders
        }

        private void AllocateButtons(DataGridView dataGridView)
        {
            DeleteTreatmentButton.Location = new Point(DeleteTreatmentButton.Location.X, dataGridView.Bottom + 10);
            SaveButton.Location = new Point(SaveButton.Location.X, DeleteTreatmentButton.Location.Y);
        }

        private void AllocateComments()
        {
            CommentsLabel.Location = new Point(CommentsLabel.Location.X, DeleteTreatmentButton.Bottom + 20);
            TreatmentCommentsTextBox.Location = new Point(CommentsLabel.Location.X, CommentsLabel.Bottom + 18);
        }

        public object TreatmentListDataSource
        {
            get => TreatmentsListBox.DataSource;
            set
            {
                TreatmentsListBox.DataSource = value;
                TreatmentsListBox.DisplayMember = nameof(Treatment.Description);
            }
        }

        public int SelectedTreatmentIndex => TreatmentsListBox.SelectedIndex;

        public event EventHandler TreatmentSelected;

        public event EventHandler DeleteLayer;

        public event EventHandler DeleteTreatment;

        public event EventHandler NewTreatment;

        public event EventHandler SaveNewTreatment;

        public event EventHandler GoBack;

        public event EventHandler SaveAllChanges;

        public event EventHandler CancelAllChanges;

        public void DisplaySelectedTreatmentDetails(Treatment treatment)
        {
            TreatmentDescriptionPanel.Visible = true;
            SaveButton.Visible = true;
            DeleteTreatmentButton.Visible = true;

            _coatingLayers.Clear();
            foreach(CoatingLayer coatingLayer in treatment.CoatingLayers)
            {
                _coatingLayers.Add(coatingLayer);
            }

            // Check if the DataGridView already exists, and reuse it
            DataGridView coatingLayersDataGridView = TreatmentDescriptionPanel.Controls
                .Find("CoatingLayersDataGridView", true)
                .FirstOrDefault() as DataGridView;

            if(coatingLayersDataGridView == null)
            {
                coatingLayersDataGridView = new DataGridView
                {
                    Location = new Point(4, 102),
                    Size = new Size(430, 100),
                    Name = "CoatingLayersDataGridView",
                    AutoGenerateColumns = false,
                    AllowUserToAddRows = true,
                    AllowUserToDeleteRows = false,
                    RowHeadersWidth = 4,
                    ColumnHeadersHeight = 50,
                    DataSource = _coatingLayers,
                    Tag = treatment,
                    ScrollBars = ScrollBars.Both
                };

                coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Product", Name = "ProductName", DataPropertyName = "ProductName", Width = 160 });
                coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Dry Film\nthickness (µm)", Name = "DryFilmThickness", DataPropertyName = "DryFilmThickness", Width = 65 });
                coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Wet Film\nthickness (µm)", Name = "WetFilmThickness", DataPropertyName = "WetFilmThickness", Width = 65 });
                coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Factual\nconsumption per m2", Name = "FactualConsumption", DataPropertyName = "FactualConsumption", Width = 80 });
                coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Units", Name = "Units", DataPropertyName = "Units", Width = 40 });
                coatingLayersDataGridView.Columns.Add(new DataGridViewButtonColumn { Text = "X", HeaderText = "", Name = "DeleteButton", UseColumnTextForButtonValue = true, Width = 20});

                coatingLayersDataGridView.CellValueChanged += CoatingLayersDataGridView_CellValueChanged;
                coatingLayersDataGridView.UserAddedRow += CoatingLayersDataGridView_UserAddedRow;
                coatingLayersDataGridView.CellContentClick += CoatingLayersDataGridView_CellContentClick;
                coatingLayersDataGridView.CellValidating += CoatingLayersDataGridView_CellValidating;
                coatingLayersDataGridView.CellEndEdit += CoatingLayersDataGridView_CellEndEdit;
                coatingLayersDataGridView.DefaultValuesNeeded += CoatingLayersDataGridView_DefaultValuesNeeded;
                coatingLayersDataGridView.DataError += CoatingLayersDataGridView_DataError;

                TreatmentDescriptionPanel.Controls.Add(coatingLayersDataGridView);
            }

            else
            {
                coatingLayersDataGridView.DataSource = null; // Reset to ensure refresh
                coatingLayersDataGridView.DataSource = _coatingLayers;
            }

            EnableDisableTreatmentDescriptionControls(false);

            ResizeDataGridViewToFitRows(coatingLayersDataGridView);

            AllocateButtons(coatingLayersDataGridView);

            AllocateComments();
        }

        private void CoatingLayersDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            // Suppress the default dialog for the DataError
            e.ThrowException = false;
        }

        private void CoatingLayersDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["ProductName"].Value = "Product Name";
            e.Row.Cells["DryFilmThickness"].Value = 1;
            e.Row.Cells["WetFilmThickness"].Value = 1;
            e.Row.Cells["FactualConsumption"].Value = 1;
            e.Row.Cells["Units"].Value = "Units";
            e.Row.Cells["DeleteButton"].Value = "X";
        }

        private void CoatingLayersDataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (sender is DataGridView dataGridView)
            {
                var columnName = dataGridView.Columns[e.ColumnIndex].DataPropertyName;

                if (columnName == "DryFilmThickness" || columnName == "WetFilmThickness" || columnName == "FactualConsumption")
                {
                    if (!double.TryParse(e.FormattedValue.ToString(), out double result) || result < 0)
                    {
                        MessageBox.Show("Value must be a valid non-negative number.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true; // Prevent invalid data
                        return;
                    }
                }
                else if (columnName == "ProductName" || columnName == "Units")
                {
                    if (string.IsNullOrWhiteSpace(e.FormattedValue.ToString()))
                    {
                        MessageBox.Show($"{columnName} cannot be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true; // Prevent empty values
                        return;
                    }
                }

                var row = dataGridView.Rows[e.RowIndex];
            }
        }

        private void CoatingLayersDataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (sender is DataGridView dataGridView && e.RowIndex >= 0)
            {
                try
                {
                    DataGridViewCell currentCell = dataGridView.CurrentCell;

                    // Ensure the treatment is valid
                    var treatment = dataGridView.Tag as Treatment;
                    _paintingSystemsManagerPresenter.ValidateTreatment(treatment);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during CellEndEdit: {ex.Message}");
                }
            }
        }

        private void CoatingLayersDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;

            // Ensure the click is not on the header row
            if (e.RowIndex >= 0 && e.ColumnIndex == 5)
            {
                // Confirm deletion
                var result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    if (TreatmentsListBox.SelectedValue == null)
                    {
                        List<int> selectedLayersIndices = GetSelectedCoatingLayerIndices();

                        // Sort the indices in descending order to avoid index issues after removal
                        selectedLayersIndices.Sort((a, b) => b.CompareTo(a));

                        foreach (int layerIndex in selectedLayersIndices)
                        {
                            if (layerIndex >= 0 && layerIndex < TempTreatment.CoatingLayers.Count)
                            {
                                TempTreatment.CoatingLayers.RemoveAt(layerIndex);
                            }
                        }

                        RefreshCoatingLayersDataGridView(TempTreatment);
                    }

                    else
                        DeleteLayer?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        private void EnableDisableTreatmentDescriptionControls(bool enabled)
        {
            foreach (Control control in TreatmentDescriptionPanel.Controls)
            {
                if (!(control is CheckBox))
                    control.Enabled = enabled;
            }
        }

        private void AdjustTextBoxSizeToFitText()
        {
            // Measure the size of the text
            Size textSize = TextRenderer.MeasureText(
                TreatmentCommentsTextBox.Text, 
                TreatmentCommentsTextBox.Font, 
                new Size(TreatmentCommentsTextBox.Width, int.MaxValue), 
                TextFormatFlags.WordBreak);

            if (textSize.Height > 160)
                textSize.Height = 160;

            // Adjust the TextBox height to fit the text
            TreatmentCommentsTextBox.Height = Math.Max(textSize.Height + 10, 60);
        }

        public void SelectAndDisplayNewlyAddedTreatment()
        {
            Control[] treatmentsListBoxes = Controls.Find("TreatmentsListBox", true);

            if (treatmentsListBoxes.Length == 0) return;

            ListBox treatmentsListBox = treatmentsListBoxes[0] as ListBox;

            treatmentsListBox.SelectedItem = TempTreatment;
        }

        private void CoatingLayersDataGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            DataGridView dataGridView = sender as DataGridView;

            // Adjust UI elements
            ResizeDataGridViewToFitRows(dataGridView);
            AllocateButtons(dataGridView);
            AllocateComments();
            
        }

        private void CoatingLayersDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (sender is DataGridView dataGridView && e.RowIndex >= 0)
            {
                CoatingLayer updatedLayer = dataGridView.Rows[e.RowIndex].DataBoundItem as CoatingLayer;

                if (TreatmentsListBox.SelectedValue == null && updatedLayer != null)
                {
                    if(TempTreatment.CoatingLayers.Count == 0)
                        TempTreatment.CoatingLayers.Add(updatedLayer);

                    else
                        TempTreatment.CoatingLayers[e.RowIndex] = updatedLayer;
                }

                else if (updatedLayer != null)
                {
                    if (!_paintingSystemsManagerPresenter.ValidateRow(dataGridView.Rows[e.RowIndex]))
                        return;

                    // For an existing treatment, sync the CoatingLayer
                    var selectedTreatment = TreatmentsListBox.SelectedValue as Treatment;
                    if (selectedTreatment != null && e.RowIndex == selectedTreatment.CoatingLayers.Count)
                    {
                        selectedTreatment.CoatingLayers.Add(updatedLayer);
                        SaveAllChanges?.Invoke(this, new EventArgs());
                    }
                }
            }
        }

        private void TreatmentCommentsTextBox_TextChanged(object sender, EventArgs e)
        {
            AdjustTextBoxSizeToFitText();
        }

        private void TreatmentsListBox_SelectedValueChanged(object sender, EventArgs e)
        {
            TempTreatment = null;
            ListBox listBox = sender as ListBox;
            Treatment selectedTreatment = listBox.SelectedValue as Treatment;

            if (selectedTreatment == null) return;

            if (selectedTreatment.Description != "New Treatment")
            {
                TreatmentSelected?.Invoke(this, EventArgs.Empty);
                EditTreatmentCheckBox.Checked = false;
            }

            else
            {
                TreatmentsListBox.SelectedIndex = -1;

                TempTreatment = new Treatment();
                TempTreatment.Type = TreatmentType;

                DisplaySelectedTreatmentDetails(TempTreatment);

                TreatmentDescription = TempTreatment.Description;
                TreatmentCleaning = TempTreatment.Cleaning;
                TreatmentComments = TempTreatment.Comments;

                EditTreatmentCheckBox.Checked = false;
                EditTreatmentCheckBox.Checked = true;

                DeleteTreatmentButton.Visible = false;
                SaveButton.Visible = true;
            }
        }

        private void DeleteTreatmentButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Are you sure you want to delete this row?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if(result == DialogResult.Yes)
            {
                DeleteTreatment?.Invoke(this, EventArgs.Empty);

                TreatmentsListBox.SelectedIndex = -1;
                TreatmentDescriptionPanel.Visible = false;
            }        
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            Panel treatmentDescriptionPanel = Controls.Find("TreatmentDescriptionPanel", true)[0] as Panel;
            DataGridView dataGridView = treatmentDescriptionPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.Cells["ProductName"].Value == "Product Name" ||
                    row.Cells["ProductName"].Value == string.Empty ||
                    row.Cells["Units"].Value == "Units" ||
                    row.Cells["Units"].Value == string.Empty)
                {
                    MessageBox.Show("Product Name and Units cannot have the default value or be empty.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            SaveAllChanges?.Invoke(this, EventArgs.Empty);
        }

        private void BackArrowButton_Click(object sender, EventArgs e)
        {
            ((Control)_compartmentWindowView).Controls.Remove(this);

            GoBack?.Invoke(this, EventArgs.Empty);
        }

        public void RefreshCoatingLayersDataGridView(Treatment treatment)
        {
            Control[] treatmentDescriptionPanels = Controls.Find("TreatmentDescriptionPanel", true);

            if (treatmentDescriptionPanels.Length == 0) return;

            Control[] coatingLayersDataGridViews = treatmentDescriptionPanels[0].Controls.Find("CoatingLayersDataGridView", true);

            if (coatingLayersDataGridViews.Length == 0) return;

            DataGridView coatingLayersDataGridView = coatingLayersDataGridViews[0] as DataGridView;

            _coatingLayers.Clear();
            foreach (CoatingLayer coatingLayer in treatment.CoatingLayers)
            {
                _coatingLayers.Add(coatingLayer);
            }

            // Resize to fit rows
            ResizeDataGridViewToFitRows(coatingLayersDataGridView);
        }

        public List<int> GetSelectedCoatingLayerIndices()
        {
            List<int> selectedIndices = new List<int>();
            Control[] treatmentDescriptionPanels = Controls.Find("TreatmentDescriptionPanel", true);

            if (treatmentDescriptionPanels.Length == 0) return selectedIndices;

            Control[] coatingLayersDataGridViews = treatmentDescriptionPanels[0].Controls.Find("CoatingLayersDataGridView", true);

            if (coatingLayersDataGridViews.Length == 0) return selectedIndices;

            DataGridView coatingLayersDataGridView = coatingLayersDataGridViews[0] as DataGridView;

            if (coatingLayersDataGridView.SelectedCells.Count == 0) return selectedIndices;

            foreach (DataGridViewCell cell in coatingLayersDataGridView.SelectedCells)
            {
                if (!selectedIndices.Contains(cell.RowIndex))
                {
                    selectedIndices.Add(cell.RowIndex);
                }
            }

            return selectedIndices;
        }

        public void RefreshBidnings(ObservableCollection<Treatment> treatments)
        {
            // Force the ListBox to refresh its bindings
            ((CurrencyManager)TreatmentsListBox.BindingContext[treatments]).Refresh();
        }

        public void ShowErrors(Dictionary<string, string> errors)
        {
            errorProvider.Clear();
            foreach (var error in errors)
            {
                // Assuming TextBoxes are named "DescriptionTextBox", "ProductNameTextBox", etc.
                Control control = this.Controls.Find(error.Key + "TextBox", true).FirstOrDefault();
                if (control != null)
                    errorProvider.SetError(control, error.Value);
                else
                {
                    // Handle errors not bound to specific controls, such as Treatments
                    MessageBox.Show(error.Value, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Implementation of ShowMessage
        public void ShowMessage(string message, string caption = "Information", MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.Information)
        {
            MessageBox.Show(message, caption, buttons, icon);
        }

        private void EditTreatmentCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;

            if(checkBox.Checked)
            {
                EnableDisableTreatmentDescriptionControls(true);
            }
            else
                EnableDisableTreatmentDescriptionControls(false);
        }
    }
}

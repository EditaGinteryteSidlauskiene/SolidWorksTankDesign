using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swdocumentmgr;
using SolidWorksTankDesign.MVP;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Presenters;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using TextBox = System.Windows.Forms.TextBox;

namespace SolidWorksTankDesign.Windows
{
    public partial class CompartmentWindowView : UserControl, ICompartmentWindowView
    {
        private  CompartmentWindowPresenter _compartmentWindowPresenter;

        private BindingSource _internalTreatmentsBindingSource;
        private Dictionary<string, PaintingAmountUnit> _paintingAmountUnitsDictionary = new Dictionary<string, PaintingAmountUnit>();

        private Control _draggingPanel = null;
        private Point _dragFromPoint;
        private int _originalIndex;
        private List<Control> _sortedPanels = new List<Control>();

        public CompartmentWindowView(string projectFolder, string mainFolderPath, string compartmentDocPath)
        {
            InitializeComponent();

            // Start initialization
            InitializeAsync(projectFolder, mainFolderPath, compartmentDocPath);
        }

        private async Task InitializeAsync(string projectFolder, string mainFolderPath, string compartmentDocPath)
        {
            // Asynchronously initialize the model and presenter
            ICompartmentConfigurationModel compartmentConfigurationModel = new CompartmentConfigurationModel(projectFolder, mainFolderPath, compartmentDocPath);
            _compartmentWindowPresenter = new CompartmentWindowPresenter(this, compartmentConfigurationModel);

            await Task.Delay(10);
            // Call the method after all async setup is complete
            ChangePreviousCompartmentRightDishedEndConfig();

            ChangeCompartmentsConfigurationLabelText();
        }

        private void ChangeCompartmentsConfigurationLabelText()
        {
            List<CompartmentConfiguration> compartmentConfigs = SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations;
            string labelText = string.Empty;

            for (int i = 0; i < compartmentConfigs.Count; i++)
            {
                string alignment = string.Empty;
                if (compartmentConfigs[i].LeftDishedEndAlignment == DishedEndAlignment.Left)
                    alignment = "(";
                else alignment = ")";

                labelText += alignment + " " + compartmentConfigs[i].Volume + " ";

                if (i == compartmentConfigs.Count - 1)
                    labelText += ")";
            }

            CompartmentsConfigurationLabel.Text = labelText;
        }

        public void UpdatePaintingSystemsComboBoxes()
        {
            Control[] compartmentPanels = Controls.Find("CompartmentPanel", true);

            if(compartmentPanels.Length > 0)
            {
                foreach (Control compartmentPanel in compartmentPanels)
                {
                    Control[] paintingSystemComboBoxes = compartmentPanel.Controls.Find("SystemComboBox", true);

                    if(paintingSystemComboBoxes.Length > 0)
                    {
                        ComboBox paintingSystemComboBox = paintingSystemComboBoxes[0] as ComboBox;

                        paintingSystemComboBox.DataSource = null;
                        paintingSystemComboBox.DataSource = _internalTreatmentsBindingSource;
                        paintingSystemComboBox.DisplayMember = "Description";
                    }
                }
            }
        }

        public void GetInternalTreatmentsBindingSource(BindingSource internalTreatmentsBindingSource)
        {
            _internalTreatmentsBindingSource = internalTreatmentsBindingSource;
        }

        public void GetPaintingAmountUnitsDictionary(Dictionary<string, PaintingAmountUnit> paintingAmountUnitsDictionary)
        {
            _paintingAmountUnitsDictionary.Clear();
            _paintingAmountUnitsDictionary = paintingAmountUnitsDictionary;
        }

        private void NewCompartmentButton_Click(object sender, EventArgs e)
        {
            NewCompartmentButtonClicked?.Invoke(this, EventArgs.Empty);

            ChangeCompartmentsConfigurationLabelText();
        }

        private void ChangeNextCompartmentLeftDishedEndConfig()
        {
            List<Control> compartmentPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            if (compartmentPanels.Count > 1)
            {
                ChangeNextCompartmentLeftDishedEndConfig(
                compartmentPanels[compartmentPanels.Count - 2].Controls.Find("RightDishedEndAlignmentComboBox", true)[0] as ComboBox, "LeftDishedEndAlignmentComboBox");

                ChangeNextCompartmentLeftDishedEndConfig(
                    compartmentPanels[compartmentPanels.Count - 2].Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as ComboBox, "ConnectionWithPreviousCompartmentComboBox");
            }
        }

        private void ChangePreviousCompartmentRightDishedEndConfig()
        {
            List<Control> compartmentPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            if (compartmentPanels.Count > 1)
            {
                for(int i =1; i < compartmentPanels.Count; i++)
                {
                    ComboBox leftDishedEndAlignment = compartmentPanels[i].Controls.Find("LeftDishedEndAlignmentComboBox", true)[0] as ComboBox;
                    ChangePreviousCompartmentRightDishedEndConfig(
                        leftDishedEndAlignment, "RightDishedEndAlignmentComboBox");

                    ComboBox connectionWithPreviousCompartment = compartmentPanels[i].Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0] as ComboBox;
                    ChangePreviousCompartmentRightDishedEndConfig(
                        connectionWithPreviousCompartment, "ConnectionWithNextCompartmentComboBox");
                }
            }
        }

        private void EnableControlsOfPreviousCompartment()
        {
            List<Control> compartmentPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");
            if(compartmentPanels.Count > 1)
            {
                Panel previousCompartmentPanel = compartmentPanels[compartmentPanels.Count - 2] as Panel;

                previousCompartmentPanel.Controls.Find("RightDishedEndAlignmentComboBox", true)[0].Enabled = true;
                previousCompartmentPanel.Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0].Enabled = true;
            }
        }

        private void CreateCompartmentPanel(CompartmentConfiguration compartmentConfiguration, int compartmentNumber)
        {
            // Add panel
            Panel newCompartmentPanel = new Panel();

            newCompartmentPanel.Tag = compartmentConfiguration;
            newCompartmentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            newCompartmentPanel.Name = "CompartmentPanel";
            newCompartmentPanel.Size = new Size(390, 450);
            newCompartmentPanel.Cursor = Cursors.Hand;
            newCompartmentPanel.MouseDown += CompartmentPanel_MouseDown;
            newCompartmentPanel.MouseMove += CompartmentPanel_MouseMove;
            newCompartmentPanel.MouseUp += CompartmentPanel_MouseUp;
            newCompartmentPanel.Click += CompartmentPanel_Click;

            List<Control> controls = UIManager.GetAscendingControls(Controls, "CompartmentPanel");
            if (controls.Count > 0)
                newCompartmentPanel.Location = new Point(19, controls.Last().Bottom + 10);
            else
                newCompartmentPanel.Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10);
            Controls.Add(newCompartmentPanel);

            // Add compartment label
            Label label = new Label();
            Binding labelBinding = new Binding("Text", compartmentConfiguration, "Name", true, DataSourceUpdateMode.OnPropertyChanged);
            label.DataBindings.Add(labelBinding);
            label.AutoSize = false;
            label.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            label.Location = new Point(4, 4);
            label.Name = "CompartmentLabel";
            label.Cursor = Cursors.Hand;
            label.Size = new Size(362, 20);
            label.Click += CompartmentLabel_Click;
            newCompartmentPanel.Controls.Add(label);

            // Add Size label
            Label sizeLabel = new Label();
            sizeLabel.AutoSize = true;
            sizeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            sizeLabel.Location = new Point(30, 38);
            sizeLabel.Text = "Size";
            newCompartmentPanel.Controls.Add(sizeLabel);

            // Add Volume label
            Label volumeLabel = new Label();
            volumeLabel.AutoSize = true;
            volumeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            volumeLabel.Location = new Point(55, 60);
            volumeLabel.Text = "Volume, m³";
            newCompartmentPanel.Controls.Add(volumeLabel);

            // Add Volume text box
            TextBox volumeTextBox = new TextBox();
            Binding volumeBinding = new Binding("Text", compartmentConfiguration, "Volume", true, DataSourceUpdateMode.OnPropertyChanged);
            volumeTextBox.DataBindings.Add(volumeBinding);
            volumeTextBox.BorderStyle = BorderStyle.FixedSingle;
            volumeTextBox.Location = new Point(221, 57);
            volumeTextBox.Name = "VolumeTextBox";
            volumeTextBox.Size = new Size(59, 20);
            volumeTextBox.TabIndex = 0;
            volumeTextBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            volumeTextBox.Tag = compartmentConfiguration.Volume;
            volumeTextBox.LostFocus += VolumeTextBox_LostFocus;
            volumeTextBox.TextChanged += VolumeTextBox_TextChanged;
            newCompartmentPanel.Controls.Add(volumeTextBox);

            // Add Length label
            Label lengthLabel = new Label();
            lengthLabel.AutoSize = true;
            lengthLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            lengthLabel.Location = new Point(55, 85);
            lengthLabel.Text = "Length, mm";
            newCompartmentPanel.Controls.Add(lengthLabel);

            // Add Length text box
            TextBox lengthTextBox = new TextBox();
            Binding lengthBinding = new Binding("Text", compartmentConfiguration, "Length", true, DataSourceUpdateMode.OnPropertyChanged);
            lengthTextBox.DataBindings.Add(lengthBinding);
            lengthTextBox.BorderStyle = BorderStyle.FixedSingle;
            lengthTextBox.Location = new Point(221, 83);
            lengthTextBox.Name = "LengthTextBox";
            lengthTextBox.Size = new Size(59, 20);
            lengthTextBox.TabIndex = 1;
            lengthTextBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            lengthTextBox.Tag = compartmentConfiguration.Length;
            lengthTextBox.TextChanged += LengthTextBox_TextChanged;
            newCompartmentPanel.Controls.Add(lengthTextBox);

            // Add Internal Surface Treatment label
            Label internalSurfaceTreatmentLabel = new Label();
            internalSurfaceTreatmentLabel.AutoSize = true;
            internalSurfaceTreatmentLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            internalSurfaceTreatmentLabel.Location = new Point(30, 122);
            internalSurfaceTreatmentLabel.Text = "Internal Surface Treatment";
            newCompartmentPanel.Controls.Add(internalSurfaceTreatmentLabel);

            // Add System Label
            Label systemLabel = new Label();
            systemLabel.AutoSize = true;
            systemLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            systemLabel.Location = new Point(55, 154);
            systemLabel.Text = "System";
            newCompartmentPanel.Controls.Add(systemLabel);

            // Add System Combo Box
            components = new System.ComponentModel.Container();
            BindingSource internalTreatmentBindingSource = new BindingSource(components);
            internalTreatmentBindingSource.DataSource = _internalTreatmentsBindingSource;

            ComboBox systemComboBox = new ComboBox();
            systemComboBox.DataBindings.Add("SelectedItem", compartmentConfiguration, "InternalSurfaceTreatment", true, DataSourceUpdateMode.OnPropertyChanged);
            systemComboBox.DataSource = internalTreatmentBindingSource;
            systemComboBox.DisplayMember = "Description";
            systemComboBox.FormattingEnabled = true;
            systemComboBox.Location = new Point(140, 150);
            systemComboBox.Name = "SystemComboBox";
            systemComboBox.Size = new Size(140, 21);
            systemComboBox.TabIndex = 2;
            systemComboBox.Tag = compartmentConfiguration;
            systemComboBox.SelectedValueChanged += SystemComboBox_SelectedValueChanged;
            newCompartmentPanel.Controls.Add(systemComboBox);

            // Add Painting Amount label
            Label paintingAmountLabel = new Label();
            paintingAmountLabel.AutoSize = true;
            paintingAmountLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            paintingAmountLabel.Location = new Point(55, 181);
            paintingAmountLabel.Text = "Amount";
            paintingAmountLabel.Visible = false;
            paintingAmountLabel.Name = "AmountLabel";
            newCompartmentPanel.Controls.Add(paintingAmountLabel);

            // Add Painting Amount panel
            Panel amountPanel = new Panel();
            amountPanel.Location = new Point(155, 177);
            amountPanel.Margin = new Padding(3, 5, 3, 3);
            amountPanel.Name = "AmountPanel";
            amountPanel.Size = new Size(125, 28);
            amountPanel.Visible = false;
            newCompartmentPanel.Controls.Add(amountPanel);

            // Add Painting Amount Text box
            TextBox amountTextBox = new TextBox();
            amountTextBox.DataBindings.Add("Text", compartmentConfiguration, "Amount", true, DataSourceUpdateMode.OnPropertyChanged, 0, "N2");
            amountTextBox.Anchor = AnchorStyles.Left;
            amountTextBox.BorderStyle = BorderStyle.FixedSingle;
            amountTextBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            amountTextBox.Location = new Point(0, 3);
            amountTextBox.Name = "AmountTextBox";
            amountTextBox.Size = new Size(79, 22);
            amountTextBox.TabIndex = 3;
            amountTextBox.Text = compartmentConfiguration.Amount.ToString();
            amountTextBox.Tag = compartmentConfiguration.Amount;
            amountTextBox.TextChanged += AmountTextBox_TextChanged;
            amountPanel.Controls.Add(amountTextBox);

            // Add Painting Amount Units combo box
            ComboBox unitsComboBox = new ComboBox();
            unitsComboBox.DataBindings.Add("SelectedValue", compartmentConfiguration, "AmountUnits", true, DataSourceUpdateMode.OnPropertyChanged);
            unitsComboBox.DataSource = new BindingSource(_paintingAmountUnitsDictionary, null);
            unitsComboBox.DisplayMember = "Key";  // Display descriptions
            unitsComboBox.ValueMember = "Value"; // Store the actual enum values 
            unitsComboBox.Anchor = AnchorStyles.Left;
            unitsComboBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            unitsComboBox.FormattingEnabled = true;
            unitsComboBox.Location = new Point(78, 2);
            unitsComboBox.Name = "AmountUnitsComboBox";
            unitsComboBox.Size = new Size(45, 24);
            unitsComboBox.TabIndex = 9;
            amountPanel.Controls.Add(unitsComboBox);

            // Add Remove Button
            Button removeCompartmentButton = new Button();
            removeCompartmentButton.Text = "X";
            removeCompartmentButton.Size = new Size(20, 20);
            removeCompartmentButton.Location = new Point(367, 0);
            removeCompartmentButton.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            removeCompartmentButton.Name = "RemoveCompartmentButton";
            removeCompartmentButton.Tag = compartmentConfiguration;
            removeCompartmentButton.Click += RemoveCompartmentButton_Click;
            newCompartmentPanel.Controls.Add(removeCompartmentButton);

            // Add dished ends config panel
            Panel dishedEndsConfigPanel = new Panel();
            dishedEndsConfigPanel.Location = new Point(29, 185);
            dishedEndsConfigPanel.Name = "DishedEndConfigPanel";
            dishedEndsConfigPanel.Size = new Size(300, 200);
            newCompartmentPanel.Controls.Add(dishedEndsConfigPanel);

            // Left dished end configuration label
            Label leftDishedEndConfigLabel = new Label();
            leftDishedEndConfigLabel.AutoSize = true;
            leftDishedEndConfigLabel.Location = new Point(0, 0);
            leftDishedEndConfigLabel.Text = "Left dished end configuration";
            leftDishedEndConfigLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(leftDishedEndConfigLabel);

            // Add left dished end alignment label
            Label leftDishedEndAlignmentLabel = new Label();
            leftDishedEndAlignmentLabel.AutoSize = true;
            leftDishedEndAlignmentLabel.Location = new Point(25, 28);
            leftDishedEndAlignmentLabel.Text = "Alignment";
            leftDishedEndAlignmentLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(leftDishedEndAlignmentLabel);

            // Add left dished end alignment combo box
            ComboBox leftDishedEndAlignmentComboBox = new ComboBox();
            leftDishedEndAlignmentComboBox.DataSource = Enum.GetValues(typeof(DishedEndAlignment));
            leftDishedEndAlignmentComboBox.DataBindings.Add("SelectedValue", compartmentConfiguration, "LeftDishedEndAlignment", true, DataSourceUpdateMode.OnPropertyChanged);
            leftDishedEndAlignmentComboBox.Location = new Point(190, 28);
            leftDishedEndAlignmentComboBox.Size = new Size(60, 20);
            leftDishedEndAlignmentComboBox.Name = "LeftDishedEndAlignmentComboBox";
            leftDishedEndAlignmentComboBox.Tag = compartmentConfiguration;
            dishedEndsConfigPanel.Controls.Add(leftDishedEndAlignmentComboBox);

            // Add Left Dished End Connection Label
            Label connectionWithPreviousCompartentLabel = new Label();
            connectionWithPreviousCompartentLabel.AutoSize = true;
            connectionWithPreviousCompartentLabel.Location = new Point(25, 55);
            connectionWithPreviousCompartentLabel.Text = "Connection with\nprevious compartment";
            connectionWithPreviousCompartentLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(connectionWithPreviousCompartentLabel);

            // Add Left Dished End Connection Combo Box
            ComboBox leftDishedEndConnectionComboBox = new ComboBox();
            leftDishedEndConnectionComboBox.FormattingEnabled = false;
            leftDishedEndConnectionComboBox.Location = new Point(190, 60);
            leftDishedEndConnectionComboBox.Name = "ConnectionWithPreviousCompartmentComboBox";
            leftDishedEndConnectionComboBox.Size = new Size(60, 21);
            leftDishedEndConnectionComboBox.DataSource = Enum.GetValues(typeof(LeftEndConnection));
            Binding leftDishedEndConnectionBinding = new Binding("SelectedValue", compartmentConfiguration, "LeftEndConnection", true, DataSourceUpdateMode.OnPropertyChanged);
            leftDishedEndConnectionComboBox.DataBindings.Add(leftDishedEndConnectionBinding);
            leftDishedEndConnectionComboBox.Tag = compartmentConfiguration;
            dishedEndsConfigPanel.Controls.Add(leftDishedEndConnectionComboBox);

            if(compartmentNumber == 0)
            {
                leftDishedEndAlignmentComboBox.Enabled = false;
                leftDishedEndConnectionComboBox.Enabled = false;
                removeCompartmentButton.Visible = false;
            }

            // Add Right Dished End Config Label
            Label rightDishedEndConfigLabel = new Label();
            rightDishedEndConfigLabel.AutoSize = true;
            rightDishedEndConfigLabel.Location = new Point(0, 100);
            rightDishedEndConfigLabel.Text = "Right dished end configuration";
            rightDishedEndConfigLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(rightDishedEndConfigLabel);

            // Add Right Dished End Alignment Label
            Label rightDishedEndsAlignmentLabel = new Label();

            rightDishedEndsAlignmentLabel.AutoSize = true;
            rightDishedEndsAlignmentLabel.Location = new Point(25, 130);
            rightDishedEndsAlignmentLabel.Text = "Alignment";
            rightDishedEndsAlignmentLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(rightDishedEndsAlignmentLabel);

            // Add Right Dished End Alignment Combo Box
            ComboBox rightDishedEndAlignmentComboBox = new ComboBox();
            rightDishedEndAlignmentComboBox.Location = new Point(190, 130);
            rightDishedEndAlignmentComboBox.Items.Add(DishedEndAlignment.Left);
            rightDishedEndAlignmentComboBox.Items.Add(DishedEndAlignment.Right);
            rightDishedEndAlignmentComboBox.SelectedItem = DishedEndAlignment.Right;
            rightDishedEndAlignmentComboBox.Size = new Size(60, 20);
            rightDishedEndAlignmentComboBox.Name = "RightDishedEndAlignmentComboBox";
            rightDishedEndAlignmentComboBox.Enabled = false;
            rightDishedEndAlignmentComboBox.SelectedValueChanged += RightDishedEndAlignmentComboBox_SelectedValueChanged;
            dishedEndsConfigPanel.Controls.Add(rightDishedEndAlignmentComboBox);

            // Add Right Dished End Connection Label
            Label rightDishedEndConnectionLabel = new Label();
            rightDishedEndConnectionLabel.AutoSize = true;
            rightDishedEndConnectionLabel.Location = new Point(25, 160);
            rightDishedEndConnectionLabel.Text = "Connection with\nnext compartment";
            rightDishedEndConnectionLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            dishedEndsConfigPanel.Controls.Add(rightDishedEndConnectionLabel);

            // Add Right Dished End Connection Combo Box
            ComboBox rightDishedEndConnectionComboBox = new ComboBox();
            rightDishedEndConnectionComboBox.FormattingEnabled = false;
            rightDishedEndConnectionComboBox.Location = new Point(190, 165);
            rightDishedEndConnectionComboBox.Name = "ConnectionWithNextCompartmentComboBox";
            rightDishedEndConnectionComboBox.Size = new Size(60, 21);
            rightDishedEndConnectionComboBox.Items.Add(LeftEndConnection.None);
            rightDishedEndConnectionComboBox.Items.Add(LeftEndConnection.Open);
            rightDishedEndConnectionComboBox.Items.Add(LeftEndConnection.Closed);
            rightDishedEndConnectionComboBox.SelectedItem = LeftEndConnection.None;
            rightDishedEndConnectionComboBox.SelectedIndex = 0;
            rightDishedEndConnectionComboBox.Enabled = false;
            rightDishedEndConnectionComboBox.SelectedValueChanged += RightDishedEndConnectionComboBox_SelectedValueChanged;
            dishedEndsConfigPanel.Controls.Add(rightDishedEndConnectionComboBox);
        }

        private void SystemComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            PaintingSystemComboBoxSelectionChanged?.Invoke(sender, EventArgs.Empty);

            ComboBox comboBox = sender as ComboBox;
            Panel currentPanel = comboBox.Parent as Panel;
            Treatment selectedInternalTreatment = comboBox.SelectedItem as Treatment;

            if (selectedInternalTreatment == null) return;

            Control[] dishedEndPanels = currentPanel.Controls.Find("DishedEndConfigPanel", true);
            if(dishedEndPanels.Length == 0) return;

            Panel dishedEndConfigPanel = dishedEndPanels[0] as Panel;

            Control[] amountLabels = currentPanel.Controls.Find("AmountLabel", true);
            Control[] amountPanels = currentPanel.Controls.Find("AmountPanel", true);

            if(amountLabels.Length > 0 && amountPanels.Length > 0)
            {
                if (selectedInternalTreatment.Description == "Without Treatment")
                {
                    Control amountLabel = amountLabels[0];
                    Control amountPanel = amountPanels[0];

                    amountLabel.Visible = false;
                    amountPanel.Visible = false;

                    dishedEndConfigPanel.Location = new Point(29, 185);
                }

                else if (selectedInternalTreatment.Description == "Manage Painting Systems")
                {
                    //List<Treatment> treatments = new List<Treatment>();

                    //foreach (Treatment treatment in _internalTreatments)
                    //{
                    //    if (treatment.Type == Treatment.TreatmentType.Internal || treatment.Type == Treatment.TreatmentType.None)
                    //        treatments.Add(treatment);
                    //}

                    //ManagePaintingSystemsWindow managePaintingSystemsWindow = new ManagePaintingSystemsWindow(this, treatments);

                    Control systemsComboBox = sender as Control;
                    CompartmentConfiguration compartmentConfiguration = systemsComboBox.Parent.Tag as CompartmentConfiguration;

                    _compartmentWindowPresenter.ShowPaintingSystemsManagerView(_compartmentWindowPresenter);
                }

                else
                {
                    Control amountLabel = amountLabels[0];
                    Control amountPanel = amountPanels[0];

                    amountLabel.Visible = true;
                    amountPanel.Visible = true;

                    dishedEndConfigPanel.Location = new Point(29, amountPanel.Bottom + 10);
                }
            }
        }

        private void UpdateRightDishedEndComboBoxesAfterRepositioning(CompartmentConfiguration referenceCompartmentConfig, Control compartmentPanelToUpdate)
        {
            ComboBox compartmentRightAlignment =
                                compartmentPanelToUpdate.Controls.Find("RightDishedEndAlignmentComboBox", true)[0] as ComboBox;

            ComboBox compartmentRightConnection =
                compartmentPanelToUpdate.Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as ComboBox;

            if(referenceCompartmentConfig ==  null)
            {
                try
                {
                    compartmentRightAlignment.Enabled = false;
                    compartmentRightAlignment.SelectedItem = DishedEndAlignment.Right;
                }
                catch { }

                try
                {
                    compartmentRightConnection.Enabled = false;
                    compartmentRightConnection.SelectedItem = LeftEndConnection.None;
                }
                catch { }
            }

            else
            {
                try
                {
                    compartmentRightAlignment.Enabled = true;
                    compartmentRightAlignment.SelectedItem = referenceCompartmentConfig.LeftDishedEndAlignment;
                }
                catch { }

                try
                {
                    compartmentRightConnection.Enabled = true;
                    compartmentRightConnection.SelectedItem = referenceCompartmentConfig.LeftEndConnection;
                }
                catch { }
            }
        }

        /// <summary>
        /// Updates affected compartments' dished ends configurations and controls.
        /// </summary>
        /// <param name="draggedCompartmentConfiguration"></param>
        /// <param name="draggingPanel"></param>
        /// <param name="changeOfIndex"></param>
        private void UpdateDishedEndsConfigsAndComboBoxes(
            CompartmentConfiguration draggedCompartmentConfiguration, Panel draggingPanel, int newIndex)
        {
            for (int i = 0; i < _sortedPanels.Count; i++)
            {
                if (_sortedPanels[i] == draggingPanel)
                {
                    // If panel dragged to the bottom
                    if (i == _sortedPanels.Count - 1)
                    {
                        // Update previous compartment's right dished end's configurations
                        UpdateRightDishedEndComboBoxesAfterRepositioning(draggedCompartmentConfiguration, _sortedPanels[i - 1]);

                        // If there are 3 and more panels
                        if (_sortedPanels.Count >= 3)
                        {
                            // Get prevopise compartment configuration
                            CompartmentConfiguration previousCompartmentConfiguration = _sortedPanels[i - 1].Tag as CompartmentConfiguration;

                            // Update compartment's that is before previous one, right dished end's configurations
                            UpdateRightDishedEndComboBoxesAfterRepositioning(previousCompartmentConfiguration, _sortedPanels[i - 2]);
                        }
                        else
                        {
                            // Update dragged first compartment's left dished end configs
                            _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(
                                _sortedPanels[0].Tag as CompartmentConfiguration, DishedEndAlignment.Left);
                            _compartmentWindowPresenter.ChangeLeftDishedEndConnection(
                                _sortedPanels[0].Tag as CompartmentConfiguration, LeftEndConnection.None);

                            EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(0);
                        }

                        if(_originalIndex == 0)
                        {
                            _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(
                                _sortedPanels[0].Tag as CompartmentConfiguration, DishedEndAlignment.Left);

                            _compartmentWindowPresenter.ChangeLeftDishedEndConnection(
                                _sortedPanels[0].Tag as CompartmentConfiguration, LeftEndConnection.None);

                            EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(0);
                        }


                        EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(_sortedPanels.Count);
                    }

                    // If the panel dragged to the top
                    else if (i == 0)
                    {
                        // Update dragged compartment's left dished end configs
                        _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(draggedCompartmentConfiguration, DishedEndAlignment.Left);
                        _compartmentWindowPresenter.ChangeLeftDishedEndConnection(draggedCompartmentConfiguration, LeftEndConnection.None);

                        // Update Combo boxes and disable
                        EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(0);

                        UpdateRightDishedEndComboBoxesAfterRepositioning(_sortedPanels[1].Tag as CompartmentConfiguration, _sortedPanels[0]);

                        // Update combo boxes of the next compartment
                        EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(1);

                        if (_sortedPanels.Count >= 3)
                            UpdateRightDishedEndComboBoxesAfterRepositioning(_sortedPanels[2].Tag as CompartmentConfiguration, _sortedPanels[1]);
                        
                        if (_sortedPanels.Count < 3 || _originalIndex == _sortedPanels.Count - 1)
                            UpdateRightDishedEndComboBoxesAfterRepositioning(null, _sortedPanels.Last());
                    }

                    else
                    {
                        // Update previous compartment's right dished end configs combo boxes
                        UpdateRightDishedEndComboBoxesAfterRepositioning(_sortedPanels[i].Tag as CompartmentConfiguration, _sortedPanels[i - 1]);

                        // Update dragged compartment's right dished end configs combo boxes
                        UpdateRightDishedEndComboBoxesAfterRepositioning(_sortedPanels[i + 1].Tag as CompartmentConfiguration, _sortedPanels[i]);

                        if (_originalIndex == 0)
                        {
                            EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(i);

                            _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(
                                _sortedPanels[0].Tag as CompartmentConfiguration, DishedEndAlignment.Left);

                            _compartmentWindowPresenter.ChangeLeftDishedEndConnection(
                                _sortedPanels[0].Tag as CompartmentConfiguration, LeftEndConnection.None);

                            EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(0);
                        }
                        else if (_originalIndex == _sortedPanels.Count - 1)
                            UpdateRightDishedEndComboBoxesAfterRepositioning(null, _sortedPanels.Last());

                        // Update right dished end combo boxes of the compartment that is previous to the original position 
                        else if (_sortedPanels.Count >= 4 && (newIndex - _originalIndex >= 2 || newIndex - _originalIndex <= -2))
                            UpdateRightDishedEndComboBoxesAfterRepositioning(
                                _sortedPanels[_originalIndex].Tag as CompartmentConfiguration, _sortedPanels[_originalIndex - 1]);
                    }
                }
            }
        }

        private void CompartmentPanel_MouseUp(object sender, MouseEventArgs e)
        {
            if (_draggingPanel != null)
            {
                CompartmentConfiguration draggedConfig =  null;
                Panel draggingPanel = _draggingPanel as Panel;

                // Determine the new index based on the dragged panel's position
                int newIndex = CalculateNewIndex(draggingPanel);

                if (newIndex != _originalIndex)
                {
                    // Move panel in the list
                    _sortedPanels.RemoveAt(_originalIndex);
                    _sortedPanels.Insert(newIndex, _draggingPanel);

                    // Update presenter/model with the new configuration order
                    draggedConfig = draggingPanel.Tag as CompartmentConfiguration;
                    _compartmentWindowPresenter.MoveCompartmentConfiguration(draggedConfig, newIndex);
                }

                if(draggedConfig != null)
                    UpdateDishedEndsConfigsAndComboBoxes(draggedConfig, draggingPanel, newIndex);

                // Rearrange panels visually
                RearrangePanels();

                _compartmentWindowPresenter.RenameCompartments();
                ChangeCompartmentsConfigurationLabelText();

                // Clear dragging state
                _draggingPanel = null;
            }
        }

        private void CompartmentPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (_draggingPanel != null)
            {
                // Update the vertical position of the dragged panel
                int newY = _draggingPanel.Top + e.Y;
                _draggingPanel.Location = new Point(_draggingPanel.Location.X, newY);

                // Optionally, force redraw during drag
                _draggingPanel.Parent.Invalidate();
            }
        }

        private void CompartmentPanel_MouseDown(object sender, MouseEventArgs e)
        {
            _draggingPanel = sender as Control;
            _sortedPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");
            _dragFromPoint = _draggingPanel.Location;

            // Store the original index of the dragged panel
            _originalIndex = _sortedPanels.IndexOf(_draggingPanel);
        }

        private int CalculateNewIndex(Panel draggedPanel)
        {
            int newIndex = 0;
            for (int i = 0; i < _sortedPanels.Count; i++)
            {
                // Skip the dragged panel itself
                if (_sortedPanels[i] == draggedPanel) continue;

                // Compare the Y-coordinate of the dragged panel with other panels
                if (draggedPanel.Top > _sortedPanels[i].Top)
                {
                    newIndex++;
                }
            }
            return newIndex;
        }

        private void RearrangePanels()
        {
            // Update the positions of all panels based on their order in the list
            for (int i = 0; i < _sortedPanels.Count; i++)
            {
                if (i == 0)
                {
                    _sortedPanels[i].Location = new Point(_sortedPanels[i].Location.X, CompartmentsConfigurationLabel.Bottom + 10);
                }

                else
                {
                    _sortedPanels[i].Location = new Point(_sortedPanels[i].Location.X, _sortedPanels[i - 1].Bottom + 10);
                }
            }
        }

        private void ChangePreviousCompartmentRightDishedEndConfig(ComboBox leftDishedEndConfigComboBox, string nameOfElementToUpdate)
        {
            Panel dishedEndsConfigPanel = leftDishedEndConfigComboBox.Parent as Panel;
            Panel compartmentPanel = dishedEndsConfigPanel.Parent as Panel;

            List<Control> sortedPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            if (sortedPanels.Count == 1) return;

            for (int i = 1; i < sortedPanels.Count; i++)
            {
                if (sortedPanels[i] == compartmentPanel)
                {
                    ComboBox comboBoxToUpdate = sortedPanels[i - 1].Controls.Find(nameOfElementToUpdate, true)[0] as ComboBox;

                    if (nameOfElementToUpdate == "ConnectionWithNextCompartmentComboBox")
                        comboBoxToUpdate.SelectedValueChanged -= RightDishedEndConnectionComboBox_SelectedValueChanged;
                    else if (nameOfElementToUpdate == "RightDishedEndAlignmentComboBox")
                        comboBoxToUpdate.SelectedValueChanged -= RightDishedEndAlignmentComboBox_SelectedValueChanged;

                    comboBoxToUpdate.SelectedItem = leftDishedEndConfigComboBox.SelectedItem;

                    if (nameOfElementToUpdate == "ConnectionWithNextCompartmentComboBox")
                        comboBoxToUpdate.SelectedValueChanged += RightDishedEndConnectionComboBox_SelectedValueChanged;
                    else if (nameOfElementToUpdate == "RightDishedEndAlignmentComboBox")
                        comboBoxToUpdate.SelectedValueChanged += RightDishedEndAlignmentComboBox_SelectedValueChanged;
                }
            }
        }

        private void ChangeNextCompartmentLeftDishedEndConfig(ComboBox rightDishedEndConfigComboBox, string nameOfElementToUpdate)
        {
            Panel dishedEndsConfigPanel = rightDishedEndConfigComboBox.Parent as Panel;
            Panel compartmentPanel = dishedEndsConfigPanel.Parent as Panel;

            List<Control> sortedPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            if (sortedPanels.Count == 1) return;

            for (int i = 0; i < sortedPanels.Count; i++)
            {
                if (sortedPanels[i] == compartmentPanel)
                {
                    ComboBox comboBoxToUpdate = sortedPanels[i+1].Controls.Find(nameOfElementToUpdate, true)[0] as ComboBox;
                    CompartmentConfiguration compartmentConfigToUpdate = comboBoxToUpdate.Tag as CompartmentConfiguration;

                    if (nameOfElementToUpdate == "ConnectionWithPreviousCompartmentComboBox")
                    {
                        comboBoxToUpdate.SelectedValueChanged -= LeftDishedEndConnectionComboBox_SelectedValueChanged;

                        _compartmentWindowPresenter.ChangeLeftDishedEndConnection(
                            compartmentConfigToUpdate, (LeftEndConnection)rightDishedEndConfigComboBox.SelectedItem);

                        comboBoxToUpdate.SelectedItem = compartmentConfigToUpdate.LeftEndConnection;

                        comboBoxToUpdate.SelectedValueChanged += LeftDishedEndConnectionComboBox_SelectedValueChanged;
                    }
                        
                    else if (nameOfElementToUpdate == "LeftDishedEndAlignmentComboBox")
                    {
                        comboBoxToUpdate.SelectedValueChanged -= LeftDishedEndAlignmentComboBox_SelectedValueChanged;

                        _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(
                            compartmentConfigToUpdate, (DishedEndAlignment)rightDishedEndConfigComboBox.SelectedItem);

                        comboBoxToUpdate.SelectedItem = compartmentConfigToUpdate.LeftDishedEndAlignment;

                        comboBoxToUpdate.SelectedValueChanged += LeftDishedEndAlignmentComboBox_SelectedValueChanged;
                    }
                }
            }
        }
        private void RightDishedEndConnectionComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox rightDishedEndConnectionComboBox = sender as ComboBox;

            ChangeNextCompartmentLeftDishedEndConfig(rightDishedEndConnectionComboBox, "ConnectionWithPreviousCompartmentComboBox");
        }

        private void RightDishedEndAlignmentComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox rightDishedEndAlignmentComboBox = sender as ComboBox;

            ChangeNextCompartmentLeftDishedEndConfig(rightDishedEndAlignmentComboBox, "LeftDishedEndAlignmentComboBox");
        }

        private void LeftDishedEndConnectionComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox leftDishedEndConnectionComboBox = sender as ComboBox;
            CompartmentConfiguration compartmentConfiguration = leftDishedEndConnectionComboBox.Tag as CompartmentConfiguration;

            ChangePreviousCompartmentRightDishedEndConfig(leftDishedEndConnectionComboBox, "ConnectionWithNextCompartmentComboBox");
            _compartmentWindowPresenter.ChangeLeftDishedEndConnection(compartmentConfiguration, (LeftEndConnection)leftDishedEndConnectionComboBox.SelectedItem);
        }

        private async void LeftDishedEndAlignmentComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            ComboBox leftDishedEndAlignmentComboBox = sender as ComboBox;
            CompartmentConfiguration compartmentConfiguration = leftDishedEndAlignmentComboBox.Tag as CompartmentConfiguration;

            ChangePreviousCompartmentRightDishedEndConfig(leftDishedEndAlignmentComboBox, "RightDishedEndAlignmentComboBox");
            _compartmentWindowPresenter.ChangeLeftDishedEndAlignment(compartmentConfiguration, (DishedEndAlignment)leftDishedEndAlignmentComboBox.SelectedItem);

            await Task.Delay(50);
            ChangeCompartmentsConfigurationLabelText();
        }

        private void RepositionFolowingPanels(List<Control> sortedControls, Point newLocation, int numberOfRemovedPanel)
        {
            for (int j = numberOfRemovedPanel + 1; j < sortedControls.Count; j++)
            {
                sortedControls[j].Location = newLocation;

                newLocation = new Point(19, sortedControls[j].Bottom + 10);
            }
        }

        /// <summary>
        /// Updates combo boxes and chages their enability. If the compartment was removed, compartmentPanelNumber = removed compartment's number.
        /// If the compartment was repositioned, compartmentPanelNumber = new position number.
        /// </summary>
        /// <param name="compartmentPanelNumber"></param>
        private void EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(int compartmentPanelNumber)
        {
            List<Control> sortedControls = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            // If there is one compartment left
            if (sortedControls.Count == 1)
            {
                sortedControls[0].Controls.Find("LeftDishedEndAlignmentComboBox", true)[0].Enabled = false;
                sortedControls[0].Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0].Enabled = false;

                ComboBox rightDishedEndAlignmentComboBox = sortedControls[0].Controls.Find("RightDishedEndAlignmentComboBox", true)[0] as ComboBox;
                rightDishedEndAlignmentComboBox.Enabled = false;
                try
                {
                    rightDishedEndAlignmentComboBox.SelectedItem = DishedEndAlignment.Right;
                }
                catch { }

                ComboBox rightEndConnectionComboBox = sortedControls[0].Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as ComboBox;
                rightEndConnectionComboBox.Enabled = false;
                try
                {
                    rightEndConnectionComboBox.SelectedItem = LeftEndConnection.None;
                }
                catch { }
                sortedControls[0].Controls.Find("RemoveCompartmentButton", true)[0].Visible = false;
            }

            // If the last compartment was removed
            else if (compartmentPanelNumber == sortedControls.Count)
            {
                sortedControls.Last().Controls.Find("LeftDishedEndAlignmentComboBox", true)[0].Enabled = true;
                sortedControls.Last().Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0].Enabled = true;

                try
                {
                    ComboBox rightDishedEndAlignmentComboBox = sortedControls.Last().Controls.Find("RightDishedEndAlignmentComboBox", true)[0] as ComboBox;
                    rightDishedEndAlignmentComboBox.Enabled = false;
                    rightDishedEndAlignmentComboBox.SelectedItem = DishedEndAlignment.Right;
                    
                }
                catch { }

                try
                {
                    ComboBox rightEndConnectionComboBox = sortedControls.Last().Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as ComboBox;
                    rightEndConnectionComboBox.Enabled = false;
                    rightEndConnectionComboBox.SelectedItem = LeftEndConnection.None;
                }
                catch { }
            }
            
            // If the first compartment was removed
            else if(compartmentPanelNumber == 0)
            {
                CompartmentConfiguration compartmentConfiguration = sortedControls[0].Tag as CompartmentConfiguration;

                ComboBox leftDishedEndAlignmentComboBox = sortedControls[0].Controls.Find("LeftDishedEndAlignmentComboBox", true)[0] as ComboBox;
                leftDishedEndAlignmentComboBox.Enabled = false;
                try
                {
                    leftDishedEndAlignmentComboBox.SelectedItem = compartmentConfiguration.LeftDishedEndAlignment;
                }
                catch {  }


                ComboBox leftEndConnectionComboBox = sortedControls[0].Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0] as ComboBox;
                leftEndConnectionComboBox.Enabled = false;

                try
                {
                    leftEndConnectionComboBox.SelectedItem = compartmentConfiguration.LeftEndConnection;
                }
                catch { }
            }

            else
            {
                ComboBox leftDishedEndAlignmentComboBox = sortedControls[compartmentPanelNumber].Controls.Find("LeftDishedEndAlignmentComboBox", true)[0] as ComboBox;
                leftDishedEndAlignmentComboBox.Enabled = true;

                ComboBox leftEndConnectionComboBox = sortedControls[compartmentPanelNumber].Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0] as ComboBox;
                leftEndConnectionComboBox.Enabled = true;
            }
        }

        private void UpdateDishedEndsConfigsAfterPanelRemove(List<Control> sortedControlsBeforeRemoval, int removedPanelPosition)
        {
            // Get previous panel's right dished end configs
            ComboBox rightDishedEndAlignment =
                sortedControlsBeforeRemoval[removedPanelPosition - 1].Controls.Find("RightDishedEndAlignmentComboBox", true)[0] as ComboBox;
            ComboBox rightDishedEndConnection =
                sortedControlsBeforeRemoval[removedPanelPosition - 1].Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as ComboBox;

            // Get next compartment's left dished end alignment and connection with previous compartment
            CompartmentConfiguration nextCompartmentConfig = 
                sortedControlsBeforeRemoval[removedPanelPosition + 1].Tag as CompartmentConfiguration;

            rightDishedEndAlignment.SelectedItem = nextCompartmentConfig.LeftDishedEndAlignment;
            rightDishedEndConnection.SelectedItem = nextCompartmentConfig.LeftEndConnection;
        }

        private void RemoveCompartmentPanel(Control panelToRemove)
        {
            List<Control> sortedControls = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            for (int i = 0; i < sortedControls.Count; i++)
            {
                // If the panel to remove is not the top panel
                if (sortedControls[i] == panelToRemove && i > 0)
                {
                    Controls.Remove(sortedControls[i]);

                    Point newLocation = new Point(19, sortedControls[i - 1].Bottom + 10);

                    RepositionFolowingPanels(sortedControls, newLocation, i);

                    RepositionLastButtons();

                    EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(i);

                    // If removed compartment is not the last
                    if (i < sortedControls.Count - 1)
                        UpdateDishedEndsConfigsAfterPanelRemove(sortedControls, i);

                    break;
                }
                // If the panel to remove is the top panel
                else if (sortedControls[i] == panelToRemove && i == 0)
                {
                    Controls.Remove(sortedControls[i]);

                    Point newLocation = new Point(19, CompartmentsConfigurationLabel.Bottom + 10);

                    RepositionFolowingPanels(sortedControls, newLocation, i);

                    RepositionLastButtons();

                    EnableOrDisableControlsAfterCompartmentPanelRemovedOrRepositioned(i);

                    break;
                }
                
            }
        }

        private void RemoveCompartmentButton_Click(object sender, EventArgs e)
        {
            Control buttonControl = sender as Control;

            Panel compartmentPanel = buttonControl.Parent as Panel;
            Label compartmentLabel = compartmentPanel.Controls.Find("CompartmentLabel", true)[0] as Label;
            
            DialogResult dialogResult = MessageBox.Show(
                $"Are you sure you want to delete {compartmentLabel.Text}?", 
                "Confirm Deletion",
                MessageBoxButtons.YesNo, 
                MessageBoxIcon.Warning);

            if(dialogResult == DialogResult.Yes)
            {
                _compartmentWindowPresenter.RemoveCompartmentConfiguration(buttonControl.Tag as CompartmentConfiguration);

                // Get Panel which the button belongs to
                Control panelControl = buttonControl.Parent;

                RemoveCompartmentPanel(panelControl);
                ChangeCompartmentsConfigurationLabelText();
            }
        }

        private void CompartmentLabel_Click(object sender, EventArgs e)
        {
            DisplayCompartmentDetails(sender);
        }

        private void CompartmentPanel_Click(object sender, EventArgs e)
        {
            DisplayCompartmentDetails(sender);
        }

        private void MinimizeAndRelocateCompartmentPanel(Control compartmentPanelToMinimize, Control panelAbove)
        {
            compartmentPanelToMinimize.Size = new Size(390, 25);

            if(panelAbove != null)
                compartmentPanelToMinimize.Location = new Point(compartmentPanelToMinimize.Location.X, panelAbove.Bottom + 10);

            else
            {
                compartmentPanelToMinimize.Location = new Point(compartmentPanelToMinimize.Location.X, CompartmentsConfigurationLabel.Bottom + 10);

            }
        }

        private void DisplayCompartmentDetails(object controlObject)
        {
            Control control = controlObject as Control;

            if (control is Label)
            {
                control = control.Parent as Panel;
            }

            List<Control> panels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            // Suspend layout for performance optimization
            this.SuspendLayout();

            try
            {
                // Step 1: Resize all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (panels[i] == control)
                    {
                        panels[i].Size = new Size(390, 450); // Enlarged panel
                        panels[i].Cursor = Cursors.Default;
                    }
                    else
                    {
                        panels[i].Size = new Size(390, 25); // Minimized panel
                        panels[i].Cursor = Cursors.Hand;
                    }
                }

                // Step 2: Reposition all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (i == 0)
                    {
                        panels[i].Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10); // First panel fixed position
                    }
                    else
                    {
                        panels[i].Location = new Point(19, panels[i - 1].Bottom + 10); // Subsequent panels
                    }
                }

                // Reposition buttons relative to the last panel
                RepositionLastButtons();
            }
            finally
            {
                // Resume layout updates
                this.ResumeLayout(true);
            }

            //Control control = controlObject as Control;

            //if (control is Label)
            //{
            //    control = control.Parent as Panel;
            //}

            //List<Control> panels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            //// Step 1: Resize all panels
            //for (int i = 0; i < panels.Count; i++)
            //{
            //    if (panels[i] == control)
            //    {
            //        panels[i].Size = new Size(390, 450); // Enlarged panel
            //        panels[i].Cursor = Cursors.Default;
            //    }
            //    else
            //    {
            //        panels[i].Size = new Size(390, 25); // Minimized panel
            //        panels[i].Cursor = Cursors.Hand;
            //    }
            //}

            //// Step 2: Reposition all panels
            //for (int i = 0; i < panels.Count; i++)
            //{
            //    if (i == 0)
            //    {
            //        panels[i].Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10); // First panel fixed position
            //    }
            //    else
            //    {
            //        panels[i].Location = new Point(19, panels[i - 1].Bottom + 10); // Subsequent panels
            //    }
            //}

            //// Reposition buttons relative to the last panel
            //RepositionLastButtons();
        }

        private void VolumeTextBox_LostFocus(object sender, EventArgs e)
        {
            TextBox volumeTextBox = sender as TextBox;

            if(volumeTextBox.Text.Replace(",", ".").EndsWith("."))
            {
                volumeTextBox.Text = volumeTextBox.Text.Split('.')[0];
            }
        }

        private void AmountTextBox_TextChanged(object sender, EventArgs e)
        {
            CheckTextBoxInputForDouble(sender);
        }

        private async void LengthTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!CheckTextBoxInputForDouble(sender)) return;

            Control control = sender as Control;
            Panel panel = control.Parent as Panel;

            await Task.Delay(10);

            Control volumeTextBox = panel.Controls.Find("VolumeTextBox", true)[0];

            volumeTextBox.TextChanged -= VolumeTextBox_TextChanged;

            _compartmentWindowPresenter.CalculateVolume(panel.Tag as CompartmentConfiguration);

            volumeTextBox.TextChanged += VolumeTextBox_TextChanged;

            ChangeCompartmentsConfigurationLabelText();
        }

        private bool CheckTextBoxInputForDouble(object sender)
        {
            bool result = true;

            TextBox textBox = sender as TextBox;
            string textBoxValue = textBox.Text;

            if(!double.TryParse(textBoxValue.Replace(" ", "").Replace(',', '.'), out _) &&
                textBoxValue != string.Empty)
            {
                double value = (double)textBox.Tag;
                textBox.Text = value.ToString();
                textBox.SelectionStart = textBox.Text.Length; // Move caret to the end
                textBox.SelectionLength = 0;

                result = false;
            }
            else if(textBoxValue == string.Empty)
                result = false;

            else if(textBoxValue.EndsWith(".") || textBoxValue.EndsWith(","))
            {
                result = false;
            }

            return result;
        }

        private async void VolumeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (!CheckTextBoxInputForDouble(sender)) return;

            Control control = sender as Control;
            Panel panel = control.Parent as Panel;

            await Task.Delay(10);

            Control lengthTextBox = panel.Controls.Find("LengthTextBox", true)[0];

            lengthTextBox.TextChanged -= LengthTextBox_TextChanged;

            _compartmentWindowPresenter.CalculateLength(panel.Tag as CompartmentConfiguration);

            lengthTextBox.TextChanged += LengthTextBox_TextChanged;

            ChangeCompartmentsConfigurationLabelText();
        }

        private void RepositionLastButtons()
        {
            NewCompartmentButton.Location = new Point(19, UIManager.GetAscendingControls(Controls, "CompartmentPanel").Last().Bottom + 10);
            BackArrowButton.Location = new Point(BackArrowButton.Location.X, NewCompartmentButton.Bottom + 50);
            ForwardArrowButton.Location = new Point(ForwardArrowButton.Location.X, NewCompartmentButton.Bottom + 50);
        }

        //private void AddCompartmentName(Panel compartmentPanel)
        //{
        //    char LatestLetter = '@';

        //    Control[] panels = Controls.Find("CompartmentPanel", true);

        //    for(int i = 0; i < panels.Length; i++)
        //    {
        //        string currentCompartmentName = panels[i].Controls.Find("CompartmentLabel", true)[0].Text;

        //        if(currentCompartmentName == string.Empty) break;

        //        string compartmentLetter = currentCompartmentName.Split(' ')[1];
        //        LatestLetter = compartmentLetter.ToCharArray()[0];
        //    }

        //    compartmentPanel.Controls.Find("CompartmentLabel", true)[0].Text = $"Compartment {(char)(LatestLetter + 1)}";
        //}

        private void AddLeftDishedEndConfigEvents()
        {
            List<Control> compartmentPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            ComboBox leftDishedAlignmentComboBox = compartmentPanels.Last().Controls.Find("LeftDishedEndAlignmentComboBox", true)[0] as ComboBox;
            ComboBox leftDishedConnectionComboBox = compartmentPanels.Last().Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0] as ComboBox;

            leftDishedAlignmentComboBox.SelectedValueChanged += LeftDishedEndAlignmentComboBox_SelectedValueChanged;
            leftDishedConnectionComboBox.SelectedValueChanged += LeftDishedEndConnectionComboBox_SelectedValueChanged;
        }

        public void AddCompartmentPanel(CompartmentConfiguration compartmentConfiguration, int compartmentNumber)
        {
            List<Control> compartmentPanels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            for (int i = 0; i < compartmentPanels.Count; i++)
            {
                if (i == 0)
                {
                    compartmentPanels[0].Controls.Find("RemoveCompartmentButton", true)[0].Visible = true;
                    MinimizeAndRelocateCompartmentPanel(compartmentPanels[i], null);
                }
                    
                else
                    MinimizeAndRelocateCompartmentPanel(compartmentPanels[i], compartmentPanels[i - 1]);
            }

            CreateCompartmentPanel(compartmentConfiguration, compartmentNumber);

            EnableControlsOfPreviousCompartment();

            RepositionLastButtons();

            AddLeftDishedEndConfigEvents();

            List<CompartmentConfiguration> compartmentConfigsFromAttribute = TankSiteDataManager.LoadTankSiteAssemblyPropertiesFromAttribute().CompartmentsConfigurations;
            if (compartmentConfigsFromAttribute.Count < SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations.Count)
                ChangeNextCompartmentLeftDishedEndConfig();
        }

        public void ClearCompartmentPanels()
        {

        }

        public void ShowPaintingSystemsManagerView(Control paintingManagerView)
        {
            Controls.Add(paintingManagerView);
            paintingManagerView.BringToFront();
        }

        public void ShowNozzleWindow(Control nozzleWindow)
        {
            Controls.Add(nozzleWindow);
            nozzleWindow.BringToFront();
        }

        public event EventHandler PaintingSystemComboBoxSelectionChanged;

        public event EventHandler NewCompartmentButtonClicked;

        public event EventHandler ForwardButtonPressed;

        public event EventHandler BackButtonPressed;

        private void ForwardArrowButton_Click(object sender, EventArgs e)
        {
            ForwardButtonPressed?.Invoke(this, EventArgs.Empty);

            BackArrowButton.Enabled = true;
        }

        private void BackArrowButton_Click(object sender, EventArgs e)
        {
            BackButtonPressed?.Invoke(this, EventArgs.Empty);
        }

      
    }
}

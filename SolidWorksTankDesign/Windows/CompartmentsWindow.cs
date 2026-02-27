using AddinWithTaskpane;
using Newtonsoft.Json;
using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using File = System.IO.File;
using TextBox = System.Windows.Forms.TextBox;

namespace SolidWorksTankDesign.Windows
{
    public partial class CompartmentsWindow : UserControl
    {
        private string _mainFolderPath;
        private string _projectFolderPath;
        private string _nozzlePositionSketchPath;
        private double _cylindricalShellVolumePerMeter;
        private double _mainDishedEndIncludeVolume;
        private double _innerDishedEndIncludeVolume;
        private double _innerDishedEndExcludeVolume;
        private List<(DishedEndAlignment dishedEndAlignment, double Volume)> _compartmentConfigurations;
        private List<double> _compartmentsLengths;

        private Control _draggingPanel = null;
        private Point _dragFromPoint;
        private int _originalIndex;
        private List<Control> _sortedPanels = new List<Control>();
        private bool _textBoxHasToBeChanged = true;

        private ObservableCollection<Treatment> _internalTreatments = new ObservableCollection<Treatment>();
        private List<CompartmentConfiguration> _compartmentsConfigurations = new List<CompartmentConfiguration>();

        public CompartmentsWindow(
            string mainFolderPath,
            string projectFolderPath,
            string nozzlePositionSketchPath,
            double cylindricalShellVolumePerMeter,
            double mainDishedEndIncludeVolume,
            double innerDishedEndIncludeVolume,
            double innerDishedEndExcludeVolume,
            double initialCompartmentLength,
            double initialCompartmentVolume)
        {
            InitializeComponent(initialCompartmentLength, initialCompartmentVolume);
            _cylindricalShellVolumePerMeter = cylindricalShellVolumePerMeter;
            _mainDishedEndIncludeVolume = mainDishedEndIncludeVolume;
            _innerDishedEndIncludeVolume = innerDishedEndIncludeVolume;
            _innerDishedEndExcludeVolume = innerDishedEndExcludeVolume;

            AmountPanel.Controls.Add(AmountTextBox);
            AmountPanel.Controls.Add(AmountUnitsComboBox);
            AmountPanel.Visible = true;

            GetInternalTreatments();

            components = new System.ComponentModel.Container();
            BindingSource internalTreatmentBindingSource = new BindingSource(components);
            internalTreatmentBindingSource.DataSource = _internalTreatments;
            SystemComboBox.DataSource = internalTreatmentBindingSource;
            SystemComboBox.DisplayMember = "Description";

            _mainFolderPath = mainFolderPath;
            _projectFolderPath = projectFolderPath;
            _nozzlePositionSketchPath = nozzlePositionSketchPath;
        }

        private void SystemComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            Panel currentPanel = comboBox.Parent as Panel;
            Treatment selectedInternalTreatment = comboBox.SelectedItem as Treatment;
            Panel dishedEndConfigPanel = currentPanel.Controls.Find("DishedEndConfigPanel", true)[0] as Panel;

            if (selectedInternalTreatment.Description == "Without Treatment")
            {
                Control amountLabel = currentPanel.Controls.Find("AmountLabel", true)[0];
                Control amountPanel = currentPanel.Controls.Find("AmountPanel", true)[0];

                ControlCollection controls = amountPanel.Controls;
                amountLabel.Visible = false;
                amountPanel.Visible = false;

                dishedEndConfigPanel.Location = new Point(29, 185);
            }

            else if (selectedInternalTreatment.Description == "Manage Painting Systems")
            {
                List<Treatment> treatments = new List<Treatment>();

                foreach (Treatment treatment in _internalTreatments)
                {
                    if (treatment.Type == Treatment.TreatmentType.Internal || treatment.Type == Treatment.TreatmentType.None)
                        treatments.Add(treatment);
                }

                ManagePaintingSystemsWindow managePaintingSystemsWindow = new ManagePaintingSystemsWindow(this, treatments);
                Controls.Add(managePaintingSystemsWindow);
                managePaintingSystemsWindow.BringToFront();
            }

            else
            {
                Control amountLabel = currentPanel.Controls.Find("AmountLabel", true)[0];
                Control amountPanel = currentPanel.Controls.Find("AmountPanel", true)[0];

                amountLabel.Visible = true;
                amountPanel.Visible = true;

                dishedEndConfigPanel.Location = new Point(29, AmountPanel.Bottom + 10);
            }
        }

        public void UpdateSystemsComboBox()
        {
            Control[] controls = Controls.Find("SystemComboBox", true);

            foreach (Control control in controls)
            {
                ComboBox comboBox = control as ComboBox;
                BindingSource bindingSource = comboBox.DataSource as BindingSource;
                bindingSource.ResetBindings(false);
            }
        }

        private void AddNewCompartmentLabel(Panel newCompartmentPanel)
        {
            Label label = new Label();
            label.AutoSize = true;
            label.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            label.Location = new Point(4, 4);
            label.Name = "CompartmentLabel";
            label.Size = new Size(112, 16);
            label.TabIndex = 0;
            label.Text = "New Compartment";
            label.Cursor = Cursors.Hand;
            label.Click += CompartmentLabel_Click;
            newCompartmentPanel.Controls.Add(label);
        }

        private void AddNewCompartmentControls(object sender)
        {
            _sortedPanels = GetAscendingControls(Controls, "CompartmentPanel");
            Control control = sender as Control;
            Panel currentCompartmentPanel;

            if (control is Label)
            {
                currentCompartmentPanel = control.Parent as Panel;
            }
            else
                currentCompartmentPanel = control as Panel;

            currentCompartmentPanel.MouseDown += CompartmentPanel_MouseDown;
            currentCompartmentPanel.MouseUp += CompartmentPanel_MouseUp;
            currentCompartmentPanel.MouseMove += CompartmentPanel_MouseMove;

            ChangeCompartmentLabelText(currentCompartmentPanel);

            AddSizeLabel(currentCompartmentPanel);

            AddVolumeLabel(currentCompartmentPanel);

            AddVolumeTextBox(currentCompartmentPanel);

            AddLengthLabel(currentCompartmentPanel);

            AddLengthTextBox(currentCompartmentPanel);

            AddInternalSurfaceTreatmentLabel(currentCompartmentPanel);

            AddSystemLabel(currentCompartmentPanel);

            AddSystemComboBox(currentCompartmentPanel);

            AddPaintingAmountLabel(currentCompartmentPanel);

            AddPaintingAmountPanel(currentCompartmentPanel);

            Panel dishedEndsConfigPanel = AddDishedEndsConfigPanel(currentCompartmentPanel);
            Control leftDishedEndAlignmentTextBox = dishedEndsConfigPanel.Controls.Find("LeftDishedEndAlignmentTextBox", true)[0];
            Control rightDishedEndAlignmentTextBox = dishedEndsConfigPanel.Controls.Find("RightDishedEndAlignmentTextBox", true)[0];

            leftDishedEndAlignmentTextBox.TextChanged += LeftDishedEndAlignmentTextBox_TextChanged;
            rightDishedEndAlignmentTextBox.TextChanged += RightDishedEndAlignmentTextBox_TextChanged;

            Panel lastPanel = AddNewCompartmentPanel();

            AddNewCompartmentLabel(lastPanel);

            ForwardButton.Location = new Point(ForwardButton.Location.X, lastPanel.Bottom + 50);
        }

        private void RefreshCompartmentPanels()
        {
            for (int i = 0; i < _sortedPanels.Count - 1; i++)
            {
                Control leftDishedEndAlignmentTextBox = _sortedPanels[i].Controls.Find("LeftDishedEndAlignmentTextBox", true)[0] as Control;
                Control connectionWithPreviousCompartmentComboBox = _sortedPanels[i].Controls.Find("ConnectionWithPreviousCompartmentComboBox", true)[0] as Control;

                Control rightDishedEndAlignmentTextBox = _sortedPanels[i].Controls.Find("RightDishedEndAlignmentTextBox", true)[0] as Control;
                Control connectionWithNextCompartmentComboBox = _sortedPanels[i].Controls.Find("ConnectionWithNextCompartmentComboBox", true)[0] as Control;

                if (i == 0)
                {
                    leftDishedEndAlignmentTextBox.Text = "(";
                    leftDishedEndAlignmentTextBox.Enabled = false;
                    ((ComboBox)connectionWithPreviousCompartmentComboBox).SelectedIndex = 0;
                    connectionWithPreviousCompartmentComboBox.Enabled = false;

                    rightDishedEndAlignmentTextBox.Enabled = true;
                    connectionWithNextCompartmentComboBox.Enabled = true;
                }
                else if (i == _sortedPanels.Count - 2)
                {
                    leftDishedEndAlignmentTextBox.Enabled = true;
                    connectionWithPreviousCompartmentComboBox.Enabled = true;

                    rightDishedEndAlignmentTextBox.Text = ")";
                    rightDishedEndAlignmentTextBox.Enabled = false;
                    ((ComboBox)connectionWithNextCompartmentComboBox).SelectedIndex = 0;
                    connectionWithNextCompartmentComboBox.Enabled = false;
                }
                else
                {
                    leftDishedEndAlignmentTextBox.Enabled = true;
                    connectionWithPreviousCompartmentComboBox.Enabled = true;

                    rightDishedEndAlignmentTextBox.Enabled = true;
                    connectionWithNextCompartmentComboBox.Enabled = true;
                }
            }
        }

        private void AddRemoveCompartmentButton()
        {
            for (int i = 0; i < _sortedPanels.Count - 1; i++)
            {
                Control[] removeCompartmentButtons = _sortedPanels[i].Controls.Find("RemoveCompartmentButton", true);
                if (removeCompartmentButtons.Length == 0)
                {
                    Button removeCompartmentButton = new Button();
                    removeCompartmentButton.Text = "X";
                    removeCompartmentButton.Size = new Size(20, 20);
                    removeCompartmentButton.Location = new Point(367, 0);
                    removeCompartmentButton.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
                    removeCompartmentButton.Name = "RemoveCompartmentButton";
                    removeCompartmentButton.Click += RemoveCompartmentButton_Click;

                    _sortedPanels[i].Controls.Add(removeCompartmentButton);
                }
                else
                    removeCompartmentButtons[0].Enabled = true;
            }
        }

        private void CompartmentPanel_MouseMove(object sender, MouseEventArgs e)
        {
            if (_draggingPanel != null)
            {
                int newY = _draggingPanel.Top + e.Y;

                // Get the last compartment panel's bottom (the one before New Compartment panel)
                if (newY >= _sortedPanels.Last().Top)
                    return;

                // Update the position of the dragged panel
                _draggingPanel.Location = new Point(_draggingPanel.Location.X, newY);// Update the position of the dragged panel
            }
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
                if(i == 0)
                {
                    _sortedPanels[i].Location = new Point(_sortedPanels[i].Location.X, 27);
                }

                else
                {
                    _sortedPanels[i].Location = new Point(_sortedPanels[i].Location.X, _sortedPanels[i - 1].Bottom + 10);
                }
            }

            RefreshCompartmentPanels();
        }

        private void CompartmentPanel_MouseUp(object sender, MouseEventArgs e)
        {
            if (_draggingPanel != null)
            {
                foreach(Control control in _sortedPanels)
                {
                    if (control == _draggingPanel)
                        control.Size = new Size(390, 450);

                    else
                        control.Size = new Size(390, 25);
                }

                // 1. Determine the new index based on the dragged panel's position
                int newIndex = CalculateNewIndex(_draggingPanel as Panel);

                // 2. Remove the panel from its original position in the list
                _sortedPanels.RemoveAt(_originalIndex);

                // 3. Insert the panel at the new index
                _sortedPanels.Insert(newIndex, _draggingPanel);

                // 4. Rearrange the panels visually
                RearrangePanels();
            }

            _draggingPanel = null;
        }

        private void CompartmentPanel_MouseDown(object sender, MouseEventArgs e)
        {
            _draggingPanel = sender as Control;
            _sortedPanels = GetAscendingControls(Controls, "CompartmentPanel");
            _dragFromPoint = _draggingPanel.Location;
            
            // Store the original index of the dragged panel
            _originalIndex = GetAscendingControls(Controls, "CompartmentPanel").IndexOf(_draggingPanel);
        }

        private void RemoveCompartmentButton_Click(object sender, EventArgs e)
        {
            Control buttonControl = sender as Control;

            // Get Panel which the button belongs to
            Control panelControl = buttonControl.Parent as Control;

            List<Control> sortedControls = GetAscendingControls(Controls, "CompartmentPanel");
            for(int i = 0; i < sortedControls.Count; i++)
            {
                if (sortedControls[i] == panelControl && i > 0)
                {
                    Controls.Remove(sortedControls[i]);

                    Point newLocation = new Point(19, sortedControls[i - 1].Bottom + 10);

                    for(int j = i +1; j < sortedControls.Count; j++)
                    {
                        sortedControls[j].Location = newLocation;

                        newLocation = new Point(19, sortedControls[j].Bottom + 10);
                    }

                    break;
                }
                else if (sortedControls[i] == panelControl && i == 0)
                {
                    Controls.Remove(sortedControls[i]);

                    Point newLocation = new Point(19, 27);

                    for (int j = i + 1; j < sortedControls.Count; j++)
                    {
                        sortedControls[j].Location = newLocation;

                        newLocation = new Point(19, sortedControls[j].Bottom + 10);
                    }
                }

                sortedControls = GetAscendingControls(Controls, "CompartmentPanel");
                if(sortedControls.Count == 2)
                {
                    sortedControls[0].Controls.Find("RemoveCompartmentButton", true)[0].Enabled = false;
                }
            }
        }

        private void CompartmentLabel_Click(object sender, EventArgs e)
        {
            string labelText;
            if (sender is Label)
            {
                Label label = sender as Label;
                labelText = label.Text;
            }
            else
            {
                Control control = sender as Control;
                Panel compartmentPanel = control as Panel;
                Label compartmentLabel = compartmentPanel.Controls.Find("CompartmentLabel", true)[0] as Label;
                labelText = compartmentLabel.Text;
            }

            DisplayCompartmentDetails(sender as Control);

            if (labelText == "New Compartment")
            {
                AddNewCompartmentControls(sender);

                // Get a list of compartment panels sorted by position from top to bottom
                _sortedPanels = GetAscendingControls(Controls, "CompartmentPanel");

                RefreshCompartmentPanels();

                AddRemoveCompartmentButton();
            }
        }

        private void CompartmentPanel_Click(object sender, EventArgs e)
        {
            Control control = sender as Control;
            Panel compartmentPanel = control as Panel;
            Label compartmentLabel = compartmentPanel.Controls.Find("CompartmentLabel", true)[0] as Label;

            DisplayCompartmentDetails(sender);

            if (compartmentLabel.Text == "New Compartment")
            {
                AddNewCompartmentControls(sender);

                // Get a list of compartment panels sorted by position from top to bottom
                _sortedPanels = GetAscendingControls(Controls, "CompartmentPanel");

                RefreshCompartmentPanels();

                AddRemoveCompartmentButton();
            }
        }

        private void LengthTextBox_TextChanged(object sender, EventArgs e)
        {
            Control control = sender as Control;
            Panel panel = control.Parent as Panel;

            TextBox lengthTextBox = panel.Controls.Find("LengthTextBox", true)[0] as TextBox;
            TextBox volumeTextBox = panel.Controls.Find("VolumeTextBox", true)[0] as TextBox;

            if (lengthTextBox.Text == string.Empty)
                volumeTextBox.Text = "0";

            else
            {
                CompartmentsManager compartmentsManager = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager;

                double newVolume = 0;
                double newLength = double.Parse(lengthTextBox.Text) / 1000;

                //Recalculate volume
                double cylindricalShellVolume = _cylindricalShellVolumePerMeter * newLength;

                // Check if there are more than 1 compartment
                if (compartmentsManager.Compartments.Count == 1)
                {
                    newVolume = cylindricalShellVolume + (_mainDishedEndIncludeVolume * 2);
                }
                else
                {
                    
                    Compartment compartment = control.Tag as Compartment;

                    int compartmentsCount = compartmentsManager.Compartments.Count;
                    for (int i = 0; i < compartmentsCount; i++)
                    {
                        // Find the count number of the compartment compartment
                        if (compartmentsManager.Compartments[i] == compartment)
                        {
                            // Check if the compartment is the first one
                            if(i == 0)
                            {
                                // Get inner dished ends volume
                                double innerDishedEndVolume =
                                    _compartmentConfigurations[0].dishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                newVolume = cylindricalShellVolume + (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                            }
                            // Check if the compartment is the last one
                            else if(i+1 == compartmentsCount)
                            {
                                // Get inner dished ends volume
                                double innerDishedEndVolume =
                                    _compartmentConfigurations.Last().dishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                newVolume = cylindricalShellVolume + (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                            }
                            else
                            {
                                double innerLeftDishedEndVolume =
                                    _compartmentConfigurations[i-1].dishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
                                double innerRightDishedEndVolume =
                                    _compartmentConfigurations[i].dishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                newVolume = cylindricalShellVolume + (innerLeftDishedEndVolume + innerRightDishedEndVolume);
                            }
                            break;
                        }
                        
                    }
                }

                volumeTextBox.TextChanged -= VolumeTextBox_TextChanged;
                volumeTextBox.Text = newVolume.ToString();
                volumeTextBox.TextChanged += VolumeTextBox_TextChanged;
            }

        }

        private void VolumeTextBox_TextChanged(object sender, EventArgs e)
        {
            Control control = sender as Control;
            Panel panel = control.Parent as Panel;

            TextBox lengthTextBox = panel.Controls.Find("LengthTextBox", true)[0] as TextBox;
            TextBox volumeTextBox = panel.Controls.Find("VolumeTextBox", true)[0] as TextBox;

            if (volumeTextBox.Text == string.Empty)
                lengthTextBox.Text = "0";

            else
            {
                CompartmentsManager compartmentsManager = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager;

                int newLength = 0;
                double requiredVolumeOfCylindrcialShell = 0;
                double newVolume = double.Parse(volumeTextBox.Text);

                //Recalculate length
                // Get volume of cylindrical shell needed
                if(compartmentsManager.Compartments.Count == 1)
                {
                    requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume * 2);
                }
                else
                {
                    Compartment compartment = control.Tag as Compartment;

                    int compartmentsCount = compartmentsManager.Compartments.Count;

                    for (int i = 0; i < compartmentsCount; i++)
                    {
                        // Find the count number of the compartment compartment
                        if (compartmentsManager.Compartments[i] == compartment)
                        {
                            // Check if the compartment is the first one
                            if (i == 0)
                            {
                                // Get inner dished ends volume
                                double innerDishedEndVolume =
                                    _compartmentConfigurations[0].dishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                            }
                            // Check if the compartment is the last one
                            else if (i + 1 == compartmentsCount)
                            {
                                // Get inner dished ends volume
                                double innerDishedEndVolume =
                                    _compartmentConfigurations.Last().dishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                requiredVolumeOfCylindrcialShell = newVolume - (_mainDishedEndIncludeVolume + innerDishedEndVolume);
                            }
                            else
                            {
                                double innerLeftDishedEndVolume =
                                    _compartmentConfigurations[i - 1].dishedEndAlignment == DishedEndAlignment.Left ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;
                                double innerRightDishedEndVolume =
                                    _compartmentConfigurations[i].dishedEndAlignment == DishedEndAlignment.Right ? _innerDishedEndIncludeVolume : _innerDishedEndExcludeVolume;

                                requiredVolumeOfCylindrcialShell = newVolume - (innerLeftDishedEndVolume + innerRightDishedEndVolume);
                            }
                            break;
                        }
                    }
                }

                // Get length of cylindrical shell
                newLength = (int)(requiredVolumeOfCylindrcialShell / _cylindricalShellVolumePerMeter * 1000);

                lengthTextBox.TextChanged -= LengthTextBox_TextChanged;
                lengthTextBox.Text = newLength.ToString();
                lengthTextBox.TextChanged += LengthTextBox_TextChanged;
            }

        }

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            CompartmentsManager compartmentsManager = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager;
            AssemblyOfDishedEnds assemblyOfDishedEnds = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfDishedEnds;

            Button button = sender as Button;
            Compartment compartment = button.Tag as Compartment;

            // Get Panel which the button belongs to
            Panel panel = button.Parent as Panel;

            // Get compartment's new length
            double.TryParse(panel.Controls.Find("LengthTextBox", false)[0].Text, out double newLength);

            compartmentsManager.ActivateDocument();
            ModelDoc2 compartmentModelDoc = compartment.GetComponent().GetModelDoc2();

            using (var compartmentDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, compartmentModelDoc))
            {
                compartment.ChangeLength(newLength / 1000);
            }

            compartmentsManager.CloseDocument();

            int i = 0;
            ModelDoc2 assemblyOfDishedEndsDoc = SolidWorksDocumentProvider._tankSiteAssembly.GetDishedEndsAssemblyComponent().GetModelDoc2();
            using(var assemblyDoc = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, assemblyOfDishedEndsDoc))
            {
                if (compartmentsManager.Compartments.Count == 1)
                    assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(newLength / 1000);

                else
                {
                    int compartmentCount = compartmentsManager.Compartments.Count;
                    // Get compartment count number
                    for (i = 0; i < compartmentCount; i++)
                    {
                        if (compartmentsManager.Compartments[i] == compartment)
                        {
                            _compartmentsLengths[i] = newLength / 1000;

                            // Change distance between dished ends
                            if (i == 0)
                            {
                                assemblyOfDishedEnds.InnerDishedEnds[0].ChangeDistance(newLength / 1000);
                            }
                            else if (i + 1 != compartmentCount)
                            {
                                assemblyOfDishedEnds.InnerDishedEnds[i].ChangeDistance(newLength / 1000);
                            }
                            else
                            {
                                assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(newLength / 1000);
                                break;
                            }

                            // Get last compartments length
                            Panel lastCompartmentPanel = Controls.Find("CompartmentPanel", true).Last() as Panel;
                            double.TryParse(lastCompartmentPanel.Controls.Find("LengthTextBox", true)[0].Text, out double lastCompartmentLength);
                            assemblyOfDishedEnds.RightDishedEnd.ChangeDistance(lastCompartmentLength / 1000);

                            break;
                        }
                    }

                    // Adjust cylindrical shell

                    // Calculate cylindrical shell's length
                    double cylindricalShellLength = 0;
                    foreach (double length in _compartmentsLengths)
                    {
                        cylindricalShellLength += length;
                    }

                    AssemblyOfCylindricalShells assemblyOfCylindricalShells = SolidWorksDocumentProvider._tankSiteAssembly._assemblyOfCylindricalShells;
                    CylindricalShell cylindricalShell = assemblyOfCylindricalShells.CylindricalShells[0];

                    ModelDoc2 assemblyOfCylindricalShellsDoc = assemblyOfCylindricalShells.ActivateDocument();
                    ModelDoc2 cylindricalShellDoc = cylindricalShell.GetComponent().GetModelDoc2();
                    using (var document = new SolidWorksDocumentWrapper(SolidWorksDocumentProvider._solidWorksApplication, cylindricalShellDoc))
                    {
                        cylindricalShell.ChangeLength(cylindricalShellLength);
                    }
                    
                    assemblyOfCylindricalShells.CloseDocument();

                    SolidWorksDocumentProvider.GetActiveDoc().EditRebuild3();
                }
            }

            UpdateCompartmentsConfigurationsList(panel, i, newLength);
        }

        private void SaveCompartmentsProperties()
        {
            // Get compartment panels
            Control[] panels = Controls.Find("CompartmentPanel", true);

            for (int i = 0; i < panels.Length; i++)
            {
                ComboBox systemComboBox = panels[i].Controls.Find("SystemComboBox", true)[0] as ComboBox;
                Treatment internalSurfaceTreatment = systemComboBox.SelectedItem as Treatment;

                double amount = 0;
                string amountUnits = string.Empty;
                Panel amountPanel = panels[i].Controls.Find("AmountPanel", true)[0] as Panel;
                if(amountPanel.Visible == true)
                {
                    string amountString = amountPanel.Controls.Find("AmountTextBox", true)[0].Text;
                    double.TryParse(amountString.Trim().Replace(',', '.'), out amount);

                    ComboBox amountUnitesComboBox = amountPanel.Controls.Find("AmountUnitsComboBox", true)[0] as ComboBox;
                    amountUnits = amountUnitesComboBox.SelectedItem.ToString();
                }

                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i].InternalSurfaceTreatment = internalSurfaceTreatment;
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i].Amount = amount;
                SolidWorksDocumentProvider._tankProperties.CompartmentsConfigurations[i].AmountUnits = amountUnits;
            }

            // Serialize properties and update Attribute
            var options = new JsonSerializerSettings { ContractResolver = new PrivatePropertyContractResolver() };
            string tankPropertiesString = JsonConvert.SerializeObject(SolidWorksDocumentProvider._tankProperties, Formatting.Indented, options);

            AttributeManager.EditAttributeParameterValue(
                SolidWorksDocumentProvider._tankSiteAssembly._tankSiteModelDoc,
                "MainEntities",
                "TankProperties",
                tankPropertiesString);
        }

        /// <summary>
        /// Defines an event handler called ForwardButton_MouseEnter. 
        /// This handler is triggered when the mouse cursor enters the area of a button (presumably a "Back" button). 
        /// When this happens, the code increases the button's size to 85 pixels wide and 33 pixels high.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ForwardButton_Click(object sender, EventArgs e)
        {
            SaveCompartmentsProperties();

            CompartmentsConfigurationWindow compartmentsConfigurationWindow = new CompartmentsConfigurationWindow(_mainFolderPath, _projectFolderPath, _nozzlePositionSketchPath);

            Controls.Add(compartmentsConfigurationWindow);
            compartmentsConfigurationWindow.BringToFront();
        }

        private void RightDishedEndAlignmentTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (!(textBox.Text == "(" || textBox.Text == ")")) return;

            Control dishedEndsConfigPanelControl = textBox.Parent;
            Control compartmentPanel = dishedEndsConfigPanelControl.Parent;

            for (int i = 0; i < _sortedPanels.Count - 1; i++)
            {
                if (_sortedPanels[i] == compartmentPanel)
                {
                    if (i == _sortedPanels.Count - 2)
                        return;
                    else
                    {
                        Panel configPanel = _sortedPanels[i + 1].Controls.Find("DishedEndConfigPanel", true)[0] as Panel;
                        Control leftDishedEndAlignmentTextBox = configPanel.Controls.Find("LeftDishedEndAlignmentTextBox", true)[0];

                        leftDishedEndAlignmentTextBox.TextChanged -= LeftDishedEndAlignmentTextBox_TextChanged;
                        leftDishedEndAlignmentTextBox.Text = textBox.Text;
                        leftDishedEndAlignmentTextBox.TextChanged += LeftDishedEndAlignmentTextBox_TextChanged;

                        return;
                    }
                }
            }
        }

        private void LeftDishedEndAlignmentTextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (!(textBox.Text == "(" || textBox.Text == ")")) return;

            Control dishedEndsConfigPanelControl = textBox.Parent;
            Control compartmentPanel = dishedEndsConfigPanelControl.Parent;

            for (int i = 0; i < _sortedPanels.Count - 1; i++)
            {
                if (_sortedPanels[i] == compartmentPanel)
                {
                    if (i == 0)
                        return;
                    else
                    {
                        Panel configPanel = _sortedPanels[i - 1].Controls.Find("DishedEndConfigPanel", true)[0] as Panel;
                        Control rightDishedEndAlignmentTextBox = configPanel.Controls.Find("RightDishedEndAlignmentTextBox", true)[0];

                        rightDishedEndAlignmentTextBox.TextChanged -= RightDishedEndAlignmentTextBox_TextChanged;
                        rightDishedEndAlignmentTextBox.Text = textBox.Text;
                        rightDishedEndAlignmentTextBox.TextChanged += RightDishedEndAlignmentTextBox_TextChanged;
                    }
                }
            }
        }
    }
}

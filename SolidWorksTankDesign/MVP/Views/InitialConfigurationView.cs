using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Presenters;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Windows
{
    public partial class InitialConfigurationView : UserControl, IInitialConfigurationView
    {
        private readonly InitialConfiguarationPresenter _initialConfigurationPresenter;
        public InitialConfigurationView(string settingsFilePath, string mainFolderPath, string compartmentDocPath)
        {
            InitializeComponent();

            _initialConfigurationPresenter = new InitialConfiguarationPresenter(this, settingsFilePath, mainFolderPath, compartmentDocPath);
        }

        public string SerialNumber 
        { 
            get => SerialNumberTextBox.Text; 
            set => SerialNumberTextBox.Text = value; 
        }
        public bool DefaultFolder { 
            get => DefaultFolderRadioButton.Checked; 
            set => DefaultFolderRadioButton.Checked = value; 
        }
        public bool SelectFolder 
        { 
            get => SelectFolderRadioButton.Checked; 
            set => SelectFolderRadioButton.Checked = value; 
        }
        public ConstructionStandard Standard 
        { 
            get => (ConstructionStandard)(ConstructionStandardComboBox.SelectedValue); 
            set => ConstructionStandardComboBox.SelectedValue = value;  
        }
        public TankType TankType 
        { 
            get => (TankType)TypeComboBox.SelectedValue; 
            set => TypeComboBox.SelectedValue = value; 
        }
        public Class Class 
        { 
            get => (Class)ClassComboBox.SelectedItem; 
            set => ClassComboBox.SelectedItem = value; 
        }
        public string Diameter 
        {
            get => DiameterTextBox.Text;
            set => DiameterTextBox.Text = value;
        }
        public string MinOperatingTemp 
        {
            get => MinOperatingTempTextBox.Text;
            set => MinOperatingTempTextBox.Text = value;
        }
        public string MaxOperatingTemp 
        {
            get => MaxOperatingTempTextBox.Text;
            set => MaxOperatingTempTextBox.Text = value;
        }
        public string MaxOperatingPressure
        {
            get => MaxOperatingPressureTextBox.Text;
            set => MaxOperatingPressureTextBox.Text = value;
        }
        public string ShellLeakTestPressure
        {
            get => ShellLeakTestPressureTextBox.Text;
            set => ShellLeakTestPressureTextBox.Text = value;
        }
        public Treatment ExternalSurfaceTreatment 
        { 
            get => (Treatment)ExternalSurfaceTreatmentComboBox.SelectedItem; 
            set => ExternalSurfaceTreatmentComboBox.SelectedItem = value; 
        }

        public void LocateRedDotNextToSerialNumberTextBox()
        {
            Panel redDotPanel =  new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "SerialNumberRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(SerialNumberTextBox.Right + 10, SerialNumberTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);

            Controls.Add(redDotPanel);
        }

        public void LocateRedDotNextToDiameterTextBox()
        {
            Panel redDotPanel = new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "DiameterRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(DiameterTextBox.Right + 10, DiameterTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);

            Controls.Add(redDotPanel);
        }

        public void LocateRedDotNextToMinOperatingTempTextBox()
        {
            Panel redDotPanel = new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "MinOperatingTempRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(MinOperatingTempTextBox.Right + 10, MinOperatingTempTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);

            Controls.Add(redDotPanel);
        }

        public void LocateRedDotNextToMaxOperatingTempTextBox()
        {
            Panel redDotPanel = new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "MaxOperatingTempRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(MaxOperatingTempTextBox.Right + 10, MaxOperatingTempTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);

            Controls.Add(redDotPanel);
        }

        public void LocateRedDotNextToMaxOperatingPressureTextBox()
        {
            Panel redDotPanel = new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "MaxOperatingPressureRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(MaxOperatingPressureTextBox.Right + 10, MaxOperatingPressureTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);


            Controls.Add(redDotPanel);
        }

        public void LocateRedDotNextToShellLeakTestPressureTextBox()
        {
            Panel redDotPanel = new Panel();
            redDotPanel.BackColor = System.Drawing.Color.Red;
            redDotPanel.Location = new Point(4, 4);
            redDotPanel.Name = "ShellLeakTestPressureRedDotPanel";
            redDotPanel.Size = new Size(10, 10);
            redDotPanel.TabIndex = 25;
            redDotPanel.Location = new Point(ShellLeakTestPressureTextBox.Right + 10, ShellLeakTestPressureTextBox.Location.Y + 4);

            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddEllipse(0, 0, redDotPanel.Width, redDotPanel.Height);
            redDotPanel.Region = new Region(path);


            Controls.Add(redDotPanel);
        }

        public void PopulateStandardsComboBox(Dictionary<string, ConstructionStandard> standardsDictionary)
        {
            ConstructionStandardComboBox.DataSource = new BindingSource(standardsDictionary, null);
            ConstructionStandardComboBox.DisplayMember = "Key";  // Display descriptions
            ConstructionStandardComboBox.ValueMember = "Value"; // Store the actual enum values 
        }

        public void PopulateTankTypeComboBox(Dictionary<string, TankType> tankTypesDictionary)
        {
            TypeComboBox.DataSource = new BindingSource(tankTypesDictionary, null);
            TypeComboBox.DisplayMember = "Key";  // Display descriptions
            TypeComboBox.ValueMember = "Value"; // Store the actual enum values
        }

        public void PopulateClassComboBox(Class[] classes)
        {
            ClassComboBox.DataSource = classes;
        }

        public void PopulateExternalTreatmentsComboBox(List<Treatment> externalTreatments)
        {
            ExternalSurfaceTreatmentComboBox.DataSource = externalTreatments;
            ExternalSurfaceTreatmentComboBox.DisplayMember = "Description";
        }

        public void ShowCompartmentsWindow(Control compartmentWindow)
        {
            Controls.Add(compartmentWindow);
            compartmentWindow.BringToFront();
        }

        public event EventHandler TextBoxGotFocused;

        public event EventHandler SaveInitialConfiguration;

        private void ShellLeakTestPressureTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("ShellLeakTestPressureRedDotPanel", true);

            if (redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void MaxOperatingPressureTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("MaxOperatingPressureRedDotPanel", true);

            if (redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void MaxOperatingTempTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("MaxOperatingTempRedDotPanel", true);

            if (redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void MinOperatingTempTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("MinOperatingTempRedDotPanel", true);

            if (redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void SerialNumberTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("SerialNumberRedDotPanel", true);

            if(redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void DiameterTextBox_GotFocus(object sender, EventArgs e)
        {
            Control[] redDots = Controls.Find("DiameterRedDotPanel", true);

            if (redDots.Length > 0)
                Controls.Remove(redDots[0]);
        }

        private void ForwardArrowButton_Click(object sender, EventArgs e)
        {
            SaveInitialConfiguration?.Invoke(this, EventArgs.Empty);
        }
    }
}

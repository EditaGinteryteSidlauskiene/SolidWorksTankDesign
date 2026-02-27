using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SolidWorksTankDesign.Windows
{
    public partial class ManagePaintingSystemsWindow : UserControl
    {
        //private CompartmentsWindow _comparmmentsWindow;

        //public ManagePaintingSystemsWindow(CompartmentsWindow compartmentsWindow, List<Treatment> treatments)
        //{
        //    _comparmmentsWindow = compartmentsWindow;
        //    InitializeComponent();

        //    AddControls(treatments);
        //}

        ///// <summary>
        ///// Creates a button for a treatment with specified properties and adds it to the TreatmentsListPanel.
        ///// </summary>
        ///// <param name="treatment"></param>
        ///// <param name="locationY"></param>
        //private void AddTreatementsButtons(Treatment treatment, int locationY)
        //{
        //    Button button = new Button();
        //    button.Text = treatment.Description;
        //    button.Font = new Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        //    button.Location = new Point(0, locationY);
        //    button.FlatStyle = FlatStyle.Flat;
        //    button.BackColor = Color.Transparent;
        //    button.AutoSize = true;
        //    button.FlatAppearance.BorderSize = 0;
        //    button.MouseHover += Button_MouseHover;
        //    button.Click += TreatmentButton_Click;  // Attach the click event handler
        //    button.Tag = treatment;                // Store the Treatment object in the Tag property
        //    TreatmentsListPanel.Controls.Add(button);
        //}

        ///// <summary>
        ///// Adjusts the height of the TreatmentsListPanel based on its contents and positions the TreatmentDetailsPanel accordingly.
        ///// </summary>
        //private void AdjustTreatmentListPanelHeight()
        //{
        //    int height = TreatmentsListPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10;
        //    if (height > 400)
        //        height = 400;  // Limit the maximum height to 400
        //    TreatmentsListPanel.Height = height;
        //    TreatmentDetailsPanel.Location = new Point(0, TreatmentsListPanel.Bottom + 20);
        //}

        ///// <summary>
        ///// Adjusts the location of the "Back" button based on the position of the TreatmentDetailsPanel.
        ///// </summary>
        //private void AdjustBackButtonLocation()
        //{
        //    Control backButtonControl = Controls.Find("BackButton", true)[0];
        //    backButtonControl.Location = new Point(50, TreatmentDetailsPanel.Bottom + 50);
        //}

        ///// <summary>
        ///// Populates the TreatmentsListPanel with buttons for each treatment and adds a "New Treatment" button if it doesn't exist.
        ///// </summary>
        ///// <param name="treatments"></param>
        //private void AddControls(List<Treatment> treatments)
        //{
        //    TreatmentsListPanel.Controls.Clear(); // Clear existing controls

        //    // Add buttons for each treatment (except the first and last, which might be special items)
        //    int lastButtonBottom = 0;
        //    for (int i = 1; i < treatments.Count - 1; i++)
        //    {
        //        Treatment treatment = treatments[i];
        //        Control[] controls = TreatmentsListPanel.Controls.Cast<Control>().ToArray();
        //        if (controls.Length > 0)
        //            lastButtonBottom = controls.Last().Bottom; // Get the bottom position of the last button

        //        AddTreatementsButtons(treatment, lastButtonBottom + 5); // Add a new button below the last one
        //    }

        //    AdjustTreatmentListPanelHeight(); // Adjust the panel height based on the added buttons

        //    // Add a "New Treatment" button if it doesn't already exist
        //    Control[] newButtonControl = TreatmentDetailsPanel.Controls.Find("NewTreatmentButton", true);
        //    if (newButtonControl.Length == 0)
        //    {
        //        AddNewTreatmentButton(0); // Add the button at the top of the panel
        //    }

        //    AdjustBackButtonLocation(); // Adjust the "Back" button location
        //}

        ///// <summary>
        ///// Removes the selected treatment from the list of treatments and updates the UI.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void DeleteTreatmentButton_Click(object sender, EventArgs e)
        //{
        //    // Get the DataGridView displaying the coating layers of the treatment to be removed.
        //    DataGridView dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;

        //    // Get the Treatment object to be removed from the DataGridView's Tag property.
        //    Treatment treatmentToRemove = dataGridView.Tag as Treatment;

        //    // Load the list of treatments from the settings data manager.
        //    List<Treatment> treatments = Helpers.SettingsDataManager.GetTreatments();

        //    // Find the treatment to remove in the loaded list.
        //    Treatment treatmentInList = treatments.Find(t => t.Description == treatmentToRemove.Description);

        //    // If the treatment is found in the list, remove it.
        //    if (treatmentInList != null)
        //    {
        //        treatments.Remove(treatmentInList);
        //    }

        //    // SaveInitialConfiguration the updated list of treatments.
        //    SettingsDataManager.SaveChangesInTreatments(treatments);

        //    // Refresh the UI to reflect the removal of the treatment.
        //    AddControls(treatments);
        //}

        ///// <summary>
        ///// Creates and configures a TextBox for entering the treatment description.
        ///// </summary>
        ///// <returns></returns>
        //private TextBox AddNewTreatmentDescriptionTextBox()
        //{
        //    TextBox newTreatmentDescription = new TextBox();
        //    newTreatmentDescription.Text = "Description";
        //    newTreatmentDescription.Location = new Point(10, 0);
        //    newTreatmentDescription.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
        //    newTreatmentDescription.Width = 300;
        //    newTreatmentDescription.Name = "Description";
        //    newTreatmentDescription.BorderStyle = BorderStyle.FixedSingle;
        //    newTreatmentDescription.GotFocus += TextBox_GotFocus; // Attach an event handler for when the TextBox gains focus
        //    TreatmentDetailsPanel.Controls.Add(newTreatmentDescription);

        //    return newTreatmentDescription;
        //}

        ///// <summary>
        ///// Creates and configures a TextBox for entering the treatment cleaning information, positioned below the description TextBox.
        ///// </summary>
        ///// <param name="newTreatmentDescriptionTextBox"></param>
        ///// <returns></returns>
        //private TextBox AddNewTreatmentCleaningTextBox(TextBox newTreatmentDescriptionTextBox)
        //{
        //    TextBox newTreatmentCleaning = new TextBox();
        //    newTreatmentCleaning.Text = "Cleaning";
        //    newTreatmentCleaning.Location = new Point(10, newTreatmentDescriptionTextBox.Bottom + 10); // Position below the description TextBox
        //    newTreatmentCleaning.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
        //    newTreatmentCleaning.Width = 300;
        //    newTreatmentCleaning.Name = "Cleaning";
        //    newTreatmentCleaning.BorderStyle = BorderStyle.FixedSingle;
        //    newTreatmentCleaning.GotFocus += TextBox_GotFocus; // Attach an event handler for when the TextBox gains focus
        //    TreatmentDetailsPanel.Controls.Add(newTreatmentCleaning);

        //    return newTreatmentCleaning;
        //}

        ///// <summary>
        ///// Creates and configures a Label for the "Coating Layers" section, positioned below the cleaning TextBox.
        ///// </summary>
        ///// <param name="newTreatmentCleaningTextBox"></param>
        ///// <returns></returns>
        //private Label AddNewTreatmentCoatingLayersLabel(TextBox newTreatmentCleaningTextBox)
        //{
        //    Label coatingLayersLabel = new Label();
        //    coatingLayersLabel.Text = "Coating Layers";
        //    coatingLayersLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
        //    coatingLayersLabel.Location = new Point(10, newTreatmentCleaningTextBox.Bottom + 10); // Position below the cleaning TextBox
        //    TreatmentDetailsPanel.Controls.Add(coatingLayersLabel);

        //    return coatingLayersLabel;
        //}

        ///// <summary>
        ///// Creates and configures a DataGridView for entering coating layer details, positioned below the coating layers label.
        ///// </summary>
        ///// <param name="newTreatmentCoatingLayersLabel"></param>
        //private void AddNewTreatmementCoatingLayersDataGridView(Label newTreatmentCoatingLayersLabel)
        //{
        //    // Create a DataGridView
        //    DataGridView coatingLayersDataGridView = new DataGridView();
        //    coatingLayersDataGridView.Name = "CoatingLayersDataGridView";

        //    // Configure DataGridView columns 
        //    coatingLayersDataGridView.AutoGenerateColumns = false;
        //    coatingLayersDataGridView.RowHeadersWidth = 5; // Set a small width for row headers

        //    // Add columns for ProductName, DryFilmThickness, etc.
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Product" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Dry Film\nthickness (µm)" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Wet Film\nthickness (µm)" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Factual\nconsumption per m2" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Units" });

        //    // Set column widths
        //    int lastColumnIndex = coatingLayersDataGridView.Columns.Count - 1;
        //    foreach (DataGridViewColumn column in coatingLayersDataGridView.Columns)
        //    {
        //        column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
        //        column.Width = 80;
        //    }
        //    coatingLayersDataGridView.Columns[lastColumnIndex].Width = 40; // Set the last column's width

        //    // Set DataGridView properties for auto-sizing
        //    coatingLayersDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        //    coatingLayersDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

        //    // Position the DataGridView below the coating layers label
        //    coatingLayersDataGridView.Location = new Point(10, newTreatmentCoatingLayersLabel.Bottom + 5);
        //    coatingLayersDataGridView.Width = 375;

        //    // Add the DataGridView to the TreatmentDetailsPanel
        //    TreatmentDetailsPanel.Controls.Add(coatingLayersDataGridView);
        //    ChangeDataGridViewHeight(coatingLayersDataGridView); // Adjust the height of the DataGridView

        //    coatingLayersDataGridView.RowsAdded += CoatingLayersDataGridView_RowsAdded; // Attach an event handler for when rows are added
        //}

        ///// <summary>
        ///// Creates and configures a TextBox for entering comments, positioned below the previous control.
        ///// </summary>
        ///// <returns></returns>
        //private TextBox AddCommentsTextBox()
        //{
        //    TextBox commentsTextBox = new TextBox();
        //    commentsTextBox.Location = new Point(10, TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10); // Position below the last control
        //    commentsTextBox.Text = "Comments";
        //    commentsTextBox.BorderStyle = BorderStyle.FixedSingle;
        //    commentsTextBox.Size = new Size(300, 80);
        //    commentsTextBox.Multiline = true;
        //    commentsTextBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
        //    commentsTextBox.Name = "Comment";
        //    commentsTextBox.GotFocus += TextBox_GotFocus; // Attach an event handler for when the TextBox gains focus
        //    TreatmentDetailsPanel.Controls.Add(commentsTextBox);

        //    return commentsTextBox;
        //}

        ///// <summary>
        ///// Creates and configures an "Add" button, positioned below the comments TextBox.
        ///// </summary>
        ///// <param name="commentsTextBox"></param>
        //private void AddAddButton(TextBox commentsTextBox)
        //{
        //    Button addButton = new Button();
        //    addButton.Text = "Add";
        //    addButton.Location = new Point(130, commentsTextBox.Bottom + 10); // Position below the comments TextBox
        //    addButton.Name = "ForwardArrow";
        //    addButton.Height = 30;
        //    addButton.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
        //    addButton.Click += AddButton_Click; // Attach the click event handler
        //    TreatmentDetailsPanel.Controls.Add(addButton);
        //}

        ///// <summary>
        ///// Handles the click event of the "New Treatment" button, clearing existing controls and adding new ones for creating a treatment.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void NewTreatmentButton_Click(object sender, EventArgs e)
        //{
        //    TreatmentDetailsPanel.Controls.Clear(); // Clear existing controls

        //    // Add controls for creating a new treatment
        //    TextBox newTreatmentDescription = AddNewTreatmentDescriptionTextBox();
        //    TextBox newTreatmentCleaning = AddNewTreatmentCleaningTextBox(newTreatmentDescription);
        //    Label coatingLayersLabel = AddNewTreatmentCoatingLayersLabel(newTreatmentCleaning);
        //    AddNewTreatmementCoatingLayersDataGridView(coatingLayersLabel);
        //    TextBox commentsTextBox = AddCommentsTextBox();
        //    AddAddButton(commentsTextBox);

        //    AdjustTreatmentDetailsPanelHeight(); // Adjust the panel height to fit the new controls
        //}


        ///// <summary>
        ///// Handles the click event of the "Add" button, adding a new treatment with the specified details.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void AddButton_Click(object sender, EventArgs e)
        //{
        //    // Validate user inputs in the treatment details panel.
        //    if (!CorrectTreatmentInputs()) { return; }

        //    // Get the values entered by the user for the new treatment.
        //    TextBox treatmentDescriptionTextBox = TreatmentDetailsPanel.Controls.Find("Description", true)[0] as TextBox;
        //    TextBox cleaningTextBox = TreatmentDetailsPanel.Controls.Find("Cleaning", true)[0] as TextBox;
        //    DataGridView coatingLayersDataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;
        //    TextBox commentsTextBox = TreatmentDetailsPanel.Controls.Find("Comment", true)[0] as TextBox;

        //    // Get the comments entered by the user, if any.
        //    string comments = string.Empty;
        //    if (commentsTextBox.Text != "Comment" && commentsTextBox.Text.Replace(" ", "") != "")
        //        comments = commentsTextBox.Text;

        //    // Extract coating layer data from the DataGridView.
        //    ObservableCollection<CoatingLayer> coatingLayers = new ObservableCollection<CoatingLayer>();
        //    for (int i = 0; i < coatingLayersDataGridView.Rows.Count - 1; i++)
        //    {
        //        DataGridViewRow row = coatingLayersDataGridView.Rows[i];

        //        // Skip empty rows.
        //        if (IsRowEmpty(row))
        //            continue;

        //        // Parse and extract coating layer properties (replace commas with periods for potential decimal values).
        //        double.TryParse(row.Cells[1].Value.ToString().Replace(',', '.'), out double dryFilmThickness);
        //        double.TryParse(row.Cells[2].Value.ToString().Replace(',', '.'), out double wetFilmThickness);
        //        double.TryParse(row.Cells[3].Value.ToString().Replace(',', '.'), out double factualConsumption);

        //        coatingLayers.Add(new CoatingLayer
        //        {
        //            ProductName = row.Cells[0].Value.ToString().Trim(),
        //            DryFilmThickness = dryFilmThickness,
        //            WetFilmThickness = wetFilmThickness,
        //            FactualConsumption = factualConsumption,
        //            Units = row.Cells[4].Value.ToString().Trim()
        //        });
        //    }

        //    // Create a new Treatment object with the extracted data.
        //    Treatment newTreatment = new Treatment
        //    {
        //        Description = treatmentDescriptionTextBox.Text,
        //        Type = Treatment.TreatmentType.Internal,
        //        Cleaning = cleaningTextBox.Text,
        //        CoatingLayers = coatingLayers,
        //        Comments = comments
        //    };

        //    // Load existing treatments and add the new treatment to the list.
        //    List<Treatment> treatments = SettingsDataManager.GetTreatments();
        //    treatments.Insert(treatments.Count - 1, newTreatment); // Insert before the last item (which might be a special item)

        //    // SaveInitialConfiguration the updated treatments list.
        //    SettingsDataManager.SaveChangesInTreatments(treatments);

        //    // Clear the treatment details panel and the treatments list panel.
        //    TreatmentDetailsPanel.Controls.Clear();
        //    TreatmentsListPanel.Controls.Clear();

        //    // Repopulate the controls with the updated treatments list.
        //    AddControls(treatments);

        //    // Adjust the height of the treatment details panel to fit the content.
        //    AdjustTreatmentDetailsPanelHeight();

        //    // Reposition the back button based on the adjusted panel height.
        //    BackButton.Location = new Point(50, TreatmentDetailsPanel.Bottom + 50);
        //}

        ///// <summary>
        ///// Helper method to check if a DataGridView row is empty.
        ///// </summary>
        ///// <param name="row"></param>
        ///// <returns></returns>
        //private bool IsRowEmpty(DataGridViewRow row)
        //{
        //    // Check if all cells in the row are empty or contain only whitespace.
        //    for (int i = 0; i < row.Cells.Count; i++)
        //    {
        //        if (!string.IsNullOrWhiteSpace(row.Cells[i].Value?.ToString()))
        //        {
        //            return false; // If any cell has a value, the row is not empty.
        //        }
        //    }
        //    return true; // All cells are empty.
        //}

        ///// <summary>
        ///// Validates the user inputs for creating a new treatment.
        ///// </summary>
        ///// <returns></returns>
        //private bool CorrectTreatmentInputs()
        //{
        //    // Get references to the input controls.
        //    TextBox treatmentDescriptionTextBox = TreatmentDetailsPanel.Controls.Find("Description", true)[0] as TextBox;
        //    TextBox cleaningTextBox = TreatmentDetailsPanel.Controls.Find("Cleaning", true)[0] as TextBox;
        //    DataGridView dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;

        //    // Load existing treatments for validation.
        //    List<Treatment> treatments = SettingsDataManager.GetTreatments();

        //    // --- Validate treatment description ---

        //    // Check if the description is empty or just the default text.
        //    if (string.IsNullOrWhiteSpace(treatmentDescriptionTextBox.Text) ||
        //        treatmentDescriptionTextBox.Text == "Description")
        //    {
        //        MessageBox.Show("Enter a description for the new treatment.");
        //        return false;
        //    }

        //    // Check for duplicate treatment descriptions (case-insensitive).
        //    foreach (Treatment treatment in treatments)
        //    {
        //        if (treatmentDescriptionTextBox.Text.Trim().Equals(treatment.Description.Trim(), StringComparison.OrdinalIgnoreCase))
        //        {
        //            MessageBox.Show($"A treatment with the description '{treatmentDescriptionTextBox.Text}' already exists.");
        //            return false;
        //        }
        //    }

        //    // --- Validate cleaning input ---

        //    // Check if the cleaning information is empty or just the default text.
        //    if (string.IsNullOrWhiteSpace(cleaningTextBox.Text) ||
        //        cleaningTextBox.Text == "Cleaning")
        //    {
        //        MessageBox.Show("Enter cleaning information for the new treatment.");
        //        return false;
        //    }

        //    // --- Validate coating layer data in the DataGridView ---

        //    for (int i = 0; i < dataGridView.Rows.Count - 1; i++)
        //    {
        //        DataGridViewRow row = dataGridView.Rows[i];

        //        // Skip empty rows.
        //        if (IsRowEmpty(row))
        //            continue;

        //        // Validate product name (cannot be empty).
        //        if (string.IsNullOrWhiteSpace(row.Cells[0].Value?.ToString()))
        //        {
        //            MessageBox.Show($"Enter the product name in row {i + 1}.");
        //            return false;
        //        }

        //        // Validate dry film thickness, wet film thickness, and factual consumption (must be numbers).
        //        if (!double.TryParse(row.Cells[1].Value?.ToString().Replace(',', '.'), out _))
        //        {
        //            MessageBox.Show($"Dry film thickness value in row {i + 1} must be a number.");
        //            return false;
        //        }
        //        if (!double.TryParse(row.Cells[2].Value?.ToString().Replace(',', '.'), out _))
        //        {
        //            MessageBox.Show($"Wet film thickness value in row {i + 1} must be a number.");
        //            return false;
        //        }
        //        if (!double.TryParse(row.Cells[3].Value?.ToString().Replace(',', '.'), out _))
        //        {
        //            MessageBox.Show($"Factual consumption per m2 value in row {i + 1} must be a number.");
        //            return false;
        //        }

        //        // Validate units (cannot be empty).
        //        if (string.IsNullOrWhiteSpace(row.Cells[4].Value?.ToString()))
        //        {
        //            MessageBox.Show($"Enter the units in row {i + 1}.");
        //            return false;
        //        }
        //    }

        //    // All inputs are valid.
        //    return true;
        //}

        ///// <summary>
        ///// Clears the text content of a TextBox when it receives focus.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void TextBox_GotFocus(object sender, EventArgs e)
        //{
        //    // Cast the sender object to a TextBox.
        //    TextBox textBox = sender as TextBox;

        //    // Clear the text content of the TextBox.
        //    textBox.Text = "";
        //}

        ///// <summary>
        ///// Adjusts the height of a DataGridView to fit its content, with a maximum height limit.
        ///// </summary>
        ///// <param name="dataGridView"></param>
        //private void ChangeDataGridViewHeight(DataGridView dataGridView)
        //{
        //    // Calculate the total height of all rows in the DataGridView.
        //    int totalRowHeight = 0;
        //    foreach (DataGridViewRow row in dataGridView.Rows)
        //    {
        //        totalRowHeight += row.Height;
        //    }

        //    // Add extra height for the column header and borders.
        //    totalRowHeight += dataGridView.ColumnHeadersHeight + 2;

        //    // Limit the maximum height of the DataGridView to 300 pixels.
        //    if (totalRowHeight > 300)
        //    {
        //        totalRowHeight = 300;
        //    }

        //    // Set the calculated height to the DataGridView.
        //    dataGridView.Height = totalRowHeight;
        //}

        ///// <summary>
        ///// Handles the RowsAdded event of the CoatingLayersDataGridView, adjusting the UI layout accordingly.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void CoatingLayersDataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        //{
        //    // Get the DataGridView that raised the event.
        //    Control dataGridViewControl = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0];
        //    DataGridView dataGridView = dataGridViewControl as DataGridView;

        //    // Adjust the height of the DataGridView to fit the new rows.
        //    ChangeDataGridViewHeight(dataGridView);

        //    // Adjust the location of the comments section to be below the DataGridView.
        //    AdjustCommentsLocation();

        //    // Adjust the location of the buttons ("Add", "Update", etc.) to be below the comments section.
        //    AdjustButtonsLocation();

        //    // Adjust the overall height of the TreatmentDetailsPanel to accommodate all controls.
        //    AdjustTreatmentDetailsPanelHeight();
        //}

        ///// <summary>
        /////  Creates and configures a TextBox to display the description of a treatment in a UI.
        ///// </summary>
        ///// <param name="treatment"></param>
        //private void AddTreatmentDescriptionLabel(Treatment treatment)
        //{
        //    // Create a new TextBox control to hold the treatment description.
        //    TextBox treatmentDescriptionLabel = new TextBox();

        //    // Set the text of the TextBox to the description of the provided treatment object.
        //    treatmentDescriptionLabel.Text = treatment.Description;

        //    // Set the font of the TextBox to Microsoft Sans Serif, 9.75pt, regular style.
        //    treatmentDescriptionLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

        //    // Position the TextBox at 10 pixels from the left and 0 pixels from the top of its container.
        //    treatmentDescriptionLabel.Location = new Point(10, 0);

        //    // Set the width of the TextBox to 375 pixels.
        //    treatmentDescriptionLabel.Width = 375;

        //    // Give the TextBox the name "TreatmentDescriptonLabel" for potential later reference.
        //    treatmentDescriptionLabel.Name = "TreatmentDescriptonLabel";

        //    // Add an event handler to the GotFocus event. This event will be triggered when the TextBox gains focus.
        //    // The TreatmentDescriptionLabel_GotFocus method (defined elsewhere) will be executed when this happens.
        //    treatmentDescriptionLabel.GotFocus += TreatmentDescriptionLabel_GotFocus;

        //    // Add the configured TextBox to the Controls collection of the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Controls.Add(treatmentDescriptionLabel);
        //}

        ///// <summary>
        ///// Creates and configures a DataGridView control to display a table of coating layers associated with a given treatment. 
        ///// It binds the DataGridView to the treatment.CoatingLayers data source, defines columns with specific headers and data properties, 
        ///// sets column widths and sizing modes, and positions the DataGridView within a TreatmentDetailsPanel. 
        ///// Additionally, it includes event handlers for cell changes, validation, and selection.
        ///// </summary>
        ///// <param name="treatment"></param>
        //private void AddTreatmentDataGridView(Treatment treatment)
        //{
        //    // Create a BindingSource for the selected Treatment's CoatingLayers
        //    BindingSource coatingLayersBindingSource = new BindingSource();
        //    coatingLayersBindingSource.DataSource = treatment.CoatingLayers;

        //    // Create a DataGridView (instead of DataGrid)
        //    DataGridView coatingLayersDataGridView = new DataGridView();
        //    coatingLayersDataGridView.DataSource = coatingLayersBindingSource;
        //    coatingLayersDataGridView.Name = "CoatingLayersDataGridView";

        //    // Configure DataGridView columns 
        //    coatingLayersDataGridView.AutoGenerateColumns = false;
        //    // Set the RowHeadersWidth property to a small value
        //    coatingLayersDataGridView.RowHeadersWidth = 5;
        //    coatingLayersDataGridView.CurrentCellDirtyStateChanged += CoatingLayersDataGridView_CurrentCellDirtyStateChanged;
        //    coatingLayersDataGridView.CellValidating += CoatingLayersDataGridView_CellValidating;

        //    // Add columns for ProductName, DryFilmThickness, etc.
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Product", DataPropertyName = "ProductName" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Dry Film\nthickness (µm)", DataPropertyName = "DryFilmThickness" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Wet Film\nthickness (µm)", DataPropertyName = "WetFilmThickness" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Factual\nconsumption per m2", DataPropertyName = "FactualConsumption" });
        //    coatingLayersDataGridView.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Units", DataPropertyName = "Units" });

        //    // Get the index of the last column
        //    int lastColumnIndex = coatingLayersDataGridView.Columns.Count - 1;

        //    foreach (DataGridViewColumn column in coatingLayersDataGridView.Columns)
        //    {
        //        column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
        //        column.Width = 80;
        //    }

        //    // Set the desired width of the last column
        //    coatingLayersDataGridView.Columns[lastColumnIndex].Width = 40;

        //    // Set ColumnHeadersHeightSizeMode to AutoSize
        //    coatingLayersDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        //    coatingLayersDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

        //    coatingLayersDataGridView.Location = new Point(15, TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10);
        //    coatingLayersDataGridView.Size = new Size(375, 50);
        //    coatingLayersDataGridView.Tag = treatment;

        //    // Add the DataGridView to the form's Controls
        //    TreatmentDetailsPanel.Controls.Add(coatingLayersDataGridView);
        //    ChangeDataGridViewHeight(coatingLayersDataGridView);
        //    coatingLayersDataGridView.ClearSelection();
        //    coatingLayersDataGridView.SelectionChanged += CoatingLayersDataGridView_SelectionChanged;
        //}

        ///// <summary>
        ///// Creates and configures a label to display the word "Comments" within a UI panel. 
        ///// It sets the text, name, position, size, and font of the label, and then adds it to the TreatmentDetailsPanel. 
        ///// The position is dynamically calculated to place the label 10 pixels below the last control already in the panel. 
        ///// This is likely used to visually separate a comments section from other elements in the UI.
        ///// </summary>
        //private void AddCommentsLabel()
        //{
        //    // Create a new Label control to display the "Comments" heading.
        //    Label commentsLabel = new Label();

        //    // Set the text of the Label to "Comments".
        //    commentsLabel.Text = "Comments";

        //    // Set the name of the Label to "CommentsLabel" for potential later reference.
        //    commentsLabel.Name = "CommentsLabel";

        //    // Position the Label 10 pixels from the left and 10 pixels below the last control in the TreatmentDetailsPanel.
        //    commentsLabel.Location = new Point(10, TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10);

        //    // Set the size of the Label to 80 pixels wide and 16 pixels high.
        //    commentsLabel.Size = new Size(80, 16);

        //    // Set the font of the Label to bold Microsoft Sans Serif, 9.75pt.
        //    commentsLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));

        //    // Add the configured Label to the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Controls.Add(commentsLabel);
        //}

        ///// <summary>
        ///// Creates and configures a RichTextBox to display comments related to a treatment. 
        ///// It sets the name, position, size, font, and border style of the RichTextBox, 
        ///// and populates it with the comments from the treatment object. It also includes an event handler 
        ///// for when the RichTextBox gains focus. Importantly, it dynamically calculates and sets the height of 
        ///// the RichTextBox based on the length of the comment text to ensure all content is visible without unnecessary scrolling.
        ///// </summary>
        ///// <param name="treatment"></param>
        //private void AddTreatmentCommentsRichText(Treatment treatment)
        //{
        //    // Create a new RichTextBox control to display and potentially edit treatment comments.
        //    RichTextBox richTextBox = new RichTextBox();

        //    // Set the name of the RichTextBox to "Comment" for potential later reference.
        //    richTextBox.Name = "Comment";

        //    // Position the RichTextBox 10 pixels from the left and 10 pixels below the last control in the TreatmentDetailsPanel.
        //    richTextBox.Location = new Point(10, TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10);

        //    // Enable multiline mode and word wrapping for the RichTextBox.
        //    richTextBox.Multiline = true;
        //    richTextBox.WordWrap = true;

        //    // Set the width of the RichTextBox to 375 pixels.
        //    richTextBox.Width = 375;

        //    // Set the font of the RichTextBox to Microsoft Sans Serif, 9.75pt, regular style.
        //    richTextBox.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

        //    // Remove the border from the RichTextBox.
        //    richTextBox.BorderStyle = BorderStyle.None;

        //    // Set the text of the RichTextBox to the comments associated with the provided treatment.
        //    richTextBox.Text = treatment.Comments;

        //    // Add an event handler to the GotFocus event. This event will be triggered when the RichTextBox gains focus.
        //    // The RichTextBox_GotFocus method (defined elsewhere) will be executed when this happens.
        //    richTextBox.GotFocus += RichTextBox_GotFocus;

        //    // Add the configured RichTextBox to the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Controls.Add(richTextBox);

        //    // Calculate the required height of the RichTextBox based on its content and formatting.
        //    Size textSize = TextRenderer.MeasureText(richTextBox.Text, richTextBox.Font,
        //                                            new Size(richTextBox.Width, int.MaxValue),
        //                                            TextFormatFlags.WordBreak);

        //    // Set the height of the RichTextBox, including padding and margins.
        //    richTextBox.Height = textSize.Height + richTextBox.Padding.Top + richTextBox.Padding.Bottom
        //                                        + richTextBox.Margin.Top + richTextBox.Margin.Bottom;
        //}

        ///// <summary>
        ///// Creates and configures a button labeled "Delete Treatment" within a UI
        ///// </summary>
        ///// <param name="newButtonControl"></param>
        //private void AddDeleteTreatmentButton(Control newButtonControl)
        //{
        //    // Create a new Button control for deleting a treatment.
        //    Button deleteTreatmentButton = new Button();

        //    // Set the text of the button to "Delete Treatment" with a newline for formatting.
        //    deleteTreatmentButton.Text = "Delete\nTreatment";

        //    // Position the button at x=280 and align its top edge with the provided newButtonControl.
        //    deleteTreatmentButton.Location = new Point(280, newButtonControl.Top);

        //    // Set the height of the button to 35 pixels.
        //    deleteTreatmentButton.Height = 35;

        //    // Add an event handler to the Click event. This event will be triggered when the button is clicked.
        //    // The DeleteTreatmentButton_Click method (defined elsewhere) will be executed when this happens.
        //    deleteTreatmentButton.Click += DeleteTreatmentButton_Click;

        //    // Set the name of the button to "DeleteTreatmentButton" for potential later reference.
        //    deleteTreatmentButton.Name = "DeleteTreatmentButton";

        //    // Add the configured button to the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Controls.Add(deleteTreatmentButton);
        //}

        ///// <summary>
        ///// Defines an event handler called TreatmentButton_Click that appears to be responsible 
        ///// for displaying details about a selected treatment in a user interface. When a treatment button is clicked,
        ///// it clears the existing content of a panel (TreatmentDetailsPanel), populates it with information about the s
        ///// elected treatment (description, coating layers, comments), and manages the positioning of "New Treatment" 
        ///// and "Delete Treatment" buttons. Finally, it adjusts the panel's height to fit the content. 
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void TreatmentButton_Click(object sender, EventArgs e)
        //{
        //    // Set the initial height of the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Height = 400;

        //    // If a DataGridView named "CoatingLayersDataGridView" exists, detach the SelectionChanged event handler.
        //    Control[] dataGrids = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true);
        //    if (dataGrids.Length > 0)
        //    {
        //        DataGridView dataGridView = dataGrids[0] as DataGridView;
        //        dataGridView.SelectionChanged -= CoatingLayersDataGridView_SelectionChanged;
        //    }

        //    // Remove all controls from the TreatmentDetailsPanel except for the "NewTreatmentButton".
        //    foreach (Control control in TreatmentDetailsPanel.Controls.OfType<Control>().ToList())
        //    {
        //        if (control is Button && control.Name == "NewTreatmentButton")
        //            continue;

        //        TreatmentDetailsPanel.Controls.Remove(control);
        //    }

        //    // Get the Treatment object associated with the clicked button.
        //    Button button = sender as Button;
        //    Treatment selectedTreatment = button.Tag as Treatment;

        //    // Add the treatment description label to the panel.
        //    AddTreatmentDescriptionLabel(selectedTreatment);

        //    // Add the DataGridView to display coating layers.
        //    AddTreatmentDataGridView(selectedTreatment);

        //    // Add the "Comments" label.
        //    AddCommentsLabel();

        //    // Add the RichTextBox for comments.
        //    AddTreatmentCommentsRichText(selectedTreatment);

        //    // Reposition the "NewTreatmentButton" or add it if it doesn't exist.
        //    Control[] newButtonControls = TreatmentDetailsPanel.Controls.Find("NewTreatmentButton", true);
        //    if (newButtonControls.Length > 0)
        //        newButtonControls[0].Location = new Point(200, TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10);
        //    else
        //        AddNewTreatmentButton(TreatmentDetailsPanel.Controls.Cast<Control>().ToArray().Last().Bottom + 10);

        //    // Add the "DeleteTreatmentButton" if it doesn't exist.
        //    Control newButtonControl = TreatmentDetailsPanel.Controls.Find("NewTreatmentButton", true)[0];
        //    Control[] deleteButtonControl = TreatmentDetailsPanel.Controls.Find("DeleteTreatmentButton", true);
        //    if (deleteButtonControl.Length == 0)
        //    {
        //        AddDeleteTreatmentButton(newButtonControl);
        //    }

        //    // Adjust the height of the TreatmentDetailsPanel to fit the content.
        //    AdjustTreatmentDetailsPanelHeight();
        //}

        ///// <summary>
        ///// This handler is triggered when a specific control, likely a TextBox named "TreatmentDescriptionLabel", 
        ///// gains focus (i.e., the user clicks on it or tabs to it). 
        ///// When this happens, the code simply calls another method called AddUpdateButton().
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void TreatmentDescriptionLabel_GotFocus(object sender, EventArgs e)
        //{
        //    // When the TreatmentDescriptionLabel (a TextBox) gains focus, call the AddUpdateButton() method.
        //    AddUpdateButton();
        //}

        ///// <summary>
        ///// This handler is triggered when a cell in a DataGridView is about to lose focus 
        ///// (e.g., the user tabs to another cell or clicks away). Its purpose is to validate
        ///// the input in specific columns that require numeric values. It checks if the cell being 
        ///// validated belongs to any of the three columns: "Dry Film thickness (µm)", "Wet Film thickness (µm)",
        ///// or "Factual consumption per m2". If it does, it attempts to parse the cell's value as a double.
        ///// If the parsing fails (indicating an invalid number), it displays an error message, cancels the cell edit, 
        ///// and disables an "Update" button. If the parsing succeeds, it enables the "Update" button. 
        ///// This ensures that only valid numeric data can be entered in these columns.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void CoatingLayersDataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        //{
        //    // Get the DataGridView that raised the event.
        //    DataGridView dataGridView = sender as DataGridView;

        //    // Get the header text of the column being validated.
        //    string columnHeader = dataGridView.Columns[e.ColumnIndex].HeaderText;

        //    // Check if the cell being validated is in any of the columns that require numeric input:
        //    // "Dry Film thickness (µm)", "Wet Film thickness (µm)", or "Factual consumption per m2"
        //    if (columnHeader == "Dry Film\nthickness (µm)" ||
        //        columnHeader == "Wet Film\nthickness (µm)" ||
        //        columnHeader == "Factual\nconsumption per m2")
        //    {
        //        // Try to parse the cell's formatted value as a double.
        //        // The out _ discards the parsed value as we only care about whether it's a valid number.
        //        if (!double.TryParse(e.FormattedValue.ToString(), out _))
        //        {
        //            // If parsing fails (not a valid number):
        //            // Display an error message to the user, including the column header for context.
        //            MessageBox.Show($"Please enter a valid number in {columnHeader}.");

        //            // Cancel the cell edit, preventing the invalid value from being committed.
        //            e.Cancel = true;

        //            // Disable the "Update" button (presumably defined elsewhere) to prevent saving invalid data.
        //            ChangeUpdateButtonEnability(false);
        //        }
        //        else
        //        {
        //            // If parsing succeeds (valid number), enable the "Update" button.
        //            ChangeUpdateButtonEnability(true);
        //        }
        //    }
        //}

        ///// <summary>
        ///// controls whether an "Update" button is enabled or disabled in a UI. 
        ///// It takes a boolean argument isEnabled to determine the desired state of the button. 
        ///// The code searches for a button named "UpdateButton" within a panel called TreatmentDetailsPanel. 
        ///// If the button is found, it sets its Enabled property to the value of isEnabled. This method is 
        ///// likely used to enable or disable the "Update" button based on the validity of data or other
        ///// conditions in the UI, preventing users from updating information when it's in an invalid state.
        ///// </summary>
        ///// <param name="isEnabled"></param>
        //private void ChangeUpdateButtonEnability(bool isEnabled)
        //{
        //    // Find the "UpdateButton" within the TreatmentDetailsPanel.
        //    Control[] controls = TreatmentDetailsPanel.Controls.Find("UpdateButton", true);

        //    // If the button is found:
        //    if (controls.Length > 0)
        //    {
        //        // Get the button and set its Enabled property to the provided isEnabled value.
        //        Button updateButton = controls[0] as Button;
        //        updateButton.Enabled = isEnabled;
        //    }
        //}

        ///// <summary>
        ///// Adds an "Update Treatment" button to a UI panel if it doesn't already exist. 
        ///// It first checks if a button with the name "UpdateButton" is present in the TreatmentDetailsPanel. 
        ///// If not, it creates a new button, sets its properties (text, position, height), 
        ///// attaches a click event handler (UpdateButton_Click), and adds it to the panel.  
        ///// Finally, it calls a method to adjust the panel's height, likely to ensure the new button is visible.
        ///// This method is probably used to dynamically add an update button when displaying treatment details, 
        ///// allowing the user to save changes to the treatment information.
        ///// </summary>
        //private void AddUpdateButton()
        //{
        //    // Check if an "UpdateButton" already exists in the TreatmentDetailsPanel.
        //    if (TreatmentDetailsPanel.Controls.Find("UpdateButton", true).Length == 0)
        //    {
        //        // If the button doesn't exist, create a new Button control.
        //        Button updateButton = new Button();

        //        // Set the text of the button to "Update Treatment".
        //        updateButton.Text = "Update Treatment";

        //        // Position the button at x=120 and align its top edge with the "NewTreatmentButton".
        //        updateButton.Location = new Point(120, TreatmentDetailsPanel.Controls.Find("NewTreatmentButton", true)[0].Top);

        //        // Set the height of the button to 35 pixels.
        //        updateButton.Height = 35;

        //        // Add an event handler to the Click event. This event will be triggered when the button is clicked.
        //        // The UpdateButton_Click method (defined elsewhere) will be executed when this happens.
        //        updateButton.Click += UpdateButton_Click;

        //        // Set the name of the button to "UpdateButton" for potential later reference.
        //        updateButton.Name = "UpdateButton";

        //        // Add the configured button to the TreatmentDetailsPanel.
        //        TreatmentDetailsPanel.Controls.Add(updateButton);
        //    }

        //    // Adjust the height of the TreatmentDetailsPanel to accommodate the new button (if added).
        //    AdjustTreatmentDetailsPanelHeight();
        //}

        ///// <summary>
        ///// This handler is triggered when a RichTextBox control gains focus 
        ///// (i.e., the user clicks on it or tabs to it). When this happens, 
        ///// the code simply calls the AddUpdateButton() method.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void RichTextBox_GotFocus(object sender, EventArgs e)
        //{
        //    // When the RichTextBox gains focus, call the AddUpdateButton() method.
        //    AddUpdateButton();
        //}

        ///// <summary>
        ///// Creates and configures a button labeled "New Treatment" within a UI. 
        ///// It sets the text, position (using the provided locationY for the vertical position), 
        ///// height, and name of the button. It also attaches a click event handler (NewTreatmentButton_Click) 
        ///// to initiate the creation of a new treatment when the button is clicked. 
        ///// Finally, it adds this button to the TreatmentDetailsPanel. 
        ///// This method is likely used to dynamically add a "New Treatment" button to the UI, allowing the user to add new treatments to the system.
        ///// </summary>
        ///// <param name="locationY"></param>
        //private void AddNewTreatmentButton(int locationY)
        //{
        //    // Create a new Button control for adding a new treatment.
        //    Button newTreatmentButton = new Button();

        //    // Set the text of the button to "New Treatment" with a newline for formatting.
        //    newTreatmentButton.Text = "New\nTreatment";

        //    // Position the button at x=200 and y=locationY (provided as an argument).
        //    newTreatmentButton.Location = new Point(200, locationY);

        //    // Set the height of the button to 35 pixels.
        //    newTreatmentButton.Height = 35;

        //    // Add an event handler to the Click event. This event will be triggered when the button is clicked.
        //    // The NewTreatmentButton_Click method (defined elsewhere) will be executed when this happens.
        //    newTreatmentButton.Click += NewTreatmentButton_Click;

        //    // Set the name of the button to "NewTreatmentButton" for potential later reference.
        //    newTreatmentButton.Name = "NewTreatmentButton";

        //    // Add the configured button to the TreatmentDetailsPanel.
        //    TreatmentDetailsPanel.Controls.Add(newTreatmentButton);
        //}

        ///// <summary>
        ///// Defines an event handler called CoatingLayersDataGridView_SelectionChanged. 
        ///// This handler is triggered when the selection in a DataGridView control (likely displaying coating layers) changes. 
        ///// Its purpose is to dynamically add a "Delete Layer" button to the UI when a layer is selected.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void CoatingLayersDataGridView_SelectionChanged(object sender, EventArgs e)
        //{
        //    // Check if a "DeleteButton" already exists in the TreatmentDetailsPanel.
        //    if (TreatmentDetailsPanel.Controls.Find("DeleteButton", true).Length == 0)
        //    {
        //        // If the button doesn't exist, create a new Button control.
        //        Button deleteButton = new Button();

        //        // Set the text of the button to "Delete Layer" with a newline for formatting.
        //        deleteButton.Text = "Delete\nLayer";

        //        // Position the button at x=40 and align its top edge with the "NewTreatmentButton".
        //        deleteButton.Location = new Point(40, TreatmentDetailsPanel.Controls.Find("NewTreatmentButton", true)[0].Top);

        //        // Set the height of the button to 35 pixels.
        //        deleteButton.Height = 35;

        //        // Add an event handler to the Click event. This event will be triggered when the button is clicked.
        //        // The DeleteButton_Click method (defined elsewhere) will be executed when this happens.
        //        deleteButton.Click += DeleteButton_Click;

        //        // Set the name of the button to "DeleteButton" for potential later reference.
        //        deleteButton.Name = "DeleteButton";

        //        // Add the configured button to the TreatmentDetailsPanel.
        //        TreatmentDetailsPanel.Controls.Add(deleteButton);
        //    }

        //    // Adjust the height of the TreatmentDetailsPanel to accommodate the new button (if added).
        //    AdjustTreatmentDetailsPanelHeight();
        //}

        ///// <summary>
        ///// Defines an event handler called DeleteButton_Click that handles the deletion of a coating layer from a DataGridView. 
        ///// When the "Delete Layer" button is clicked, it removes the selected coating layer from the associated 
        ///// Treatment object's CoatingLayers collection. It then updates the list of treatments stored in the 
        ///// SettingsDataManager and refreshes the DataGridView to reflect the changes. Finally, it adjusts the height 
        ///// of the DataGridView to fit the remaining content. This handler essentially manages the deletion and UI update
        ///// process when a user removes a coating layer from a treatment.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void DeleteButton_Click(object sender, EventArgs e)
        //{
        //    // Get the DataGridView that displays the coating layers.
        //    DataGridView dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;

        //    // Get the index of the selected row in the DataGridView.
        //    int rowIndex = dataGridView.SelectedCells[0].RowIndex;

        //    // Get the Treatment object associated with the DataGridView.
        //    Treatment treatment = dataGridView.Tag as Treatment;

        //    // Remove the coating layer at the selected row index from the Treatment's CoatingLayers collection.
        //    treatment.CoatingLayers.RemoveAt(rowIndex);

        //    // Get the list of all treatments.
        //    List<Treatment> treatments = SettingsDataManager.GetTreatments();
        //    if (treatments != null)
        //    {
        //        // Find the treatment in the list that matches the current treatment (based on description).
        //        for (int i = 0; i < treatments.Count; i++)
        //        {
        //            if (treatments[i].Description == treatment.Description)
        //            {
        //                // Replace the old treatment with the updated treatment in the list.
        //                treatments[i] = treatment;
        //                break;
        //            }
        //        }
        //    }

        //    // SaveInitialConfiguration the updated list of treatments.
        //    SettingsDataManager.SaveChangesInTreatments(treatments);

        //    // Update the DataGridView's data source with the modified CoatingLayers collection.
        //    BindingSource coatingLayersBindingSource = new BindingSource();
        //    coatingLayersBindingSource.DataSource = treatment.CoatingLayers;
        //    dataGridView.DataSource = coatingLayersBindingSource;

        //    // Refresh the DataGridView to reflect the changes.
        //    dataGridView.Refresh();

        //    // Adjust the height of the DataGridView to fit the remaining content.
        //    ChangeDataGridViewHeight(dataGridView);
        //}

        ///// <summary>
        ///// Defines a method called AdjustTreatmentDetailsPanelHeight that dynamically adjusts the
        ///// height of a panel named TreatmentDetailsPanel to fit its content. It iterates through all the 
        ///// controls within the panel, identifies the highest and lowest controls based on their vertical position,
        ///// and calculates the necessary panel height to accommodate them. It adds a 20-pixel buffer to the calculated height, 
        ///// likely for visual spacing. Additionally, it repositions a "BackButton" control (presumably located outside the panel)
        ///// below the adjusted panel. This method ensures that the panel is always sized correctly to display all its content
        ///// without unnecessary empty space and that the "BackButton" remains positioned appropriately below the panel.
        ///// </summary>
        //private void AdjustTreatmentDetailsPanelHeight()
        //{
        //    // Get all controls within the TreatmentDetailsPanel.
        //    Control[] panelControls = TreatmentDetailsPanel.Controls.Cast<Control>().ToArray();

        //    // Initialize variables to track the highest and lowest controls in the panel.
        //    Control highestControl = panelControls[0];
        //    Control lowestControl = panelControls[0];

        //    // Iterate through the controls to find the highest and lowest ones based on their position.
        //    for (int i = 1; i < panelControls.Length; i++)
        //    {
        //        if (panelControls[i].Top < highestControl.Top)
        //            highestControl = panelControls[i];
        //        else if (panelControls[i].Bottom > lowestControl.Bottom)
        //            lowestControl = panelControls[i];
        //    }

        //    // Calculate the required height of the panel based on the highest and lowest controls, adding a 20-pixel buffer.
        //    TreatmentDetailsPanel.Height = lowestControl.Bottom - highestControl.Top + 20;

        //    // Reposition the "BackButton" (presumably outside the panel) below the adjusted TreatmentDetailsPanel.
        //    Control backButtonControl = Controls.Find("BackButton", true)[0];
        //    backButtonControl.Location = new Point(50, TreatmentDetailsPanel.Bottom + 50);
        //}

        ///// <summary>
        ///// Defines a method called AdjustCommentsLocation that adjusts the position of a comments section within a UI panel.
        ///// It aims to place the comments section, consisting of a "CommentsLabel" (if it exists) and a "Comment" RichTextBox, below a DataGridView control.
        ///// </summary>
        //private void AdjustCommentsLocation()
        //{
        //    // Get the DataGridView and the "Comment" RichTextBox from the TreatmentDetailsPanel.
        //    Control dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as Control;
        //    Control comment = TreatmentDetailsPanel.Controls.Find("Comment", true)[0] as Control;

        //    // Try to find the "CommentsLabel".
        //    Control[] commentLabels = TreatmentDetailsPanel.Controls.Find("CommentsLabel", true);
        //    if (commentLabels.Length > 0)
        //    {
        //        // If the "CommentsLabel" exists, position it below the DataGridView.
        //        commentLabels[0].Location = new Point(10, dataGridView.Bottom + 10);

        //        // Position the "Comment" RichTextBox below the "CommentsLabel".
        //        comment.Location = new Point(10, commentLabels[0].Bottom + 5);
        //    }
        //    else
        //    {
        //        // If the "CommentsLabel" doesn't exist, position the "Comment" RichTextBox directly below the DataGridView.
        //        comment.Location = new Point(10, dataGridView.Bottom + 10);
        //    }
        //}

        ///// <summary>
        ///// Adjusts the vertical position of all buttons within a UI panel named TreatmentDetailsPanel. 
        ///// It first retrieves the "Comment" RichTextBox from the panel. Then, it iterates through all the controls in the panel.
        ///// If a control is a Button, it repositions that button 10 pixels below the bottom of the "Comment" 
        ///// RichTextBox, while keeping its original horizontal position. This method ensures that all buttons in the
        ///// panel are consistently positioned below the comments section, maintaining a clean and organized UI layout.
        ///// </summary>
        //private void AdjustButtonsLocation()
        //{
        //    // Get the "Comment" RichTextBox from the TreatmentDetailsPanel.
        //    Control comment = TreatmentDetailsPanel.Controls.Find("Comment", true)[0] as Control;

        //    // Iterate through all controls in the TreatmentDetailsPanel.
        //    foreach (Control control in TreatmentDetailsPanel.Controls)
        //    {
        //        // If the control is a Button:
        //        if (control is Button)
        //        {
        //            // Reposition the button to be 10 pixels below the "Comment" RichTextBox, maintaining its original X-coordinate.
        //            control.Location = new Point(control.Location.X, comment.Bottom + 10);
        //        }
        //    }
        //}

        ///// <summary>
        ///// This handler is triggered when the content of a cell in the "CoatingLayersDataGridView" changes. 
        ///// It seems to be responsible for dynamically adjusting the layout of UI elements within the TreatmentDetailsPanel
        ///// in response to changes in the DataGridView.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void CoatingLayersDataGridView_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        //{
        //    // Get the "CoatingLayersDataGridView" from the TreatmentDetailsPanel.
        //    DataGridView dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;

        //    // Adjust the height of the DataGridView to fit its content.
        //    ChangeDataGridViewHeight(dataGridView);

        //    // Adjust the position of the comments section ("CommentsLabel" and "Comment" RichTextBox).
        //    AdjustCommentsLocation();

        //    // Adjust the position of all buttons in the panel.
        //    AdjustButtonsLocation();

        //    // Adjust the overall height of the TreatmentDetailsPanel to fit all its content.
        //    AdjustTreatmentDetailsPanelHeight();

        //    // Add the "Update" button to the panel if it doesn't already exist.
        //    AddUpdateButton();
        //}

        ///// <summary>
        ///// Defines an event handler called UpdateButton_Click that handles the process of updating a treatment's information.
        ///// When the "Update Treatment" button is clicked, it retrieves the updated treatment description and comments from the UI,
        ///// finds the corresponding treatment in a list of treatments, updates its properties, saves the changes, and refreshes the
        ///// UI to reflect the updated information. This handler essentially manages the entire update workflow for treatment details.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void UpdateButton_Click(object sender, EventArgs e)
        //{
        //    // Get the TextBox containing the treatment description, the DataGridView, and the Treatment object to update.
        //    TextBox definitionTextBox = TreatmentDetailsPanel.Controls.Find("TreatmentDescriptonLabel", true)[0] as TextBox;
        //    DataGridView dataGridView = TreatmentDetailsPanel.Controls.Find("CoatingLayersDataGridView", true)[0] as DataGridView;
        //    Treatment treatmentToUpdate = dataGridView.Tag as Treatment;

        //    // Validate if a new description is provided.
        //    if (definitionTextBox.Text == null || definitionTextBox.Text.Replace(" ", "") == "")
        //    {
        //        MessageBox.Show("Enter a new description.");
        //        return;
        //    }

        //    // Update the Comments property of the treatment with the text from the "Comment" RichTextBox.
        //    treatmentToUpdate.Comments = TreatmentDetailsPanel.Controls.Find("Comment", true)[0].Text;

        //    // Get the list of all treatments.
        //    List<Treatment> treatments = SettingsDataManager.GetTreatments();

        //    // Find the treatment in the list that matches the current treatment (based on description) and update it.
        //    for (int i = 1; i < treatments.Count - 1; i++)
        //    {
        //        if (treatments[i].Description == treatmentToUpdate.Description)
        //        {
        //            treatments[i] = treatmentToUpdate;
        //            break;
        //        }
        //    }

        //    // Update the Description property of the treatment with the text from the TextBox.
        //    treatmentToUpdate.Description = definitionTextBox.Text;

        //    // SaveInitialConfiguration the updated list of treatments.
        //    SettingsDataManager.SaveChangesInTreatments(treatments);

        //    // Refresh the UI with the updated treatments (likely by re-adding controls to the panel).
        //    AddControls(treatments);
        //}

        ///// <summary>
        ///// his handler is triggered when the mouse hovers over a button. It modifies the appearance of
        ///// the button by setting its FlatStyle to Flat (giving it a modern, borderless look) and making its background transparent. 
        ///// This likely creates a visual effect where the button blends in more with its background when the mouse hovers over it.
        ///// This is a common technique used to provide visual feedback to the user and enhance the interactivity of the UI.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void Button_MouseHover(object sender, EventArgs e)
        //{
        //    // Cast the sender object to a Button.
        //    Button button = sender as Button;

        //    // Set the button's FlatStyle to Flat for a more modern look.
        //    button.FlatStyle = FlatStyle.Flat;

        //    // Make the button's background transparent.
        //    button.BackColor = Color.Transparent;
        //}

        ///// <summary>
        ///// Defines an event handler called BackButton_MouseLeave. 
        ///// This handler is triggered when the mouse cursor leaves the area of a button (presumably a "Back" button). 
        ///// When this happens, the code sets the button's size to 75 pixels wide and 23 pixels high.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void BackButton_MouseLeave(object sender, EventArgs e)
        //{
        //    // Cast the sender object to a Button (assuming this event handler is attached to a Button).
        //    Button button = sender as Button;

        //    // Set the button's size to 75 pixels wide and 23 pixels high.
        //    button.Size = new Size(75, 23);
        //}

        ///// <summary>
        ///// Defines an event handler called BackButton_MouseEnter. 
        ///// This handler is triggered when the mouse cursor enters the area of a button (presumably a "Back" button). 
        ///// When this happens, the code increases the button's size to 85 pixels wide and 33 pixels high.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void BackButton_MouseEnter(object sender, EventArgs e)
        //{
        //    // Cast the sender object to a Button (assuming this event handler is attached to a Button).
        //    Button button = sender as Button;

        //    // Set the button's size to 85 pixels wide and 33 pixels high.
        //    button.Size = new Size(85, 33);
        //}

        ///// <summary>
        ///// Defines an event handler called BackButton_Click. This handler is executed when a button, presumably a "Back" button, is clicked. It seems to be responsible for navigating back to a previous window or updating the content of another window.
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void BackButton_Click(object sender, EventArgs e)
        //{
        //    // Hide the current window/form.
        //    this.Visible = false;

        //    // Call the GetInternalTreatments() method of an object named _comparmmentsWindow.
        //    _comparmmentsWindow.GetInternalTreatments();

        //    // Call the UpdateSystemsComboBox() method of the same _comparmmentsWindow object.
        //    _comparmmentsWindow.UpdateSystemsComboBox();
        //}
    }
}

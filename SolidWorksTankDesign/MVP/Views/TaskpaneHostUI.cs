using AddinWithTaskpane;
using MVP;
using SolidWorks.Interop.sldworks;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SolidWorksTankDesign
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl, ITaskpaneHostUI
    {
        private TaskpaneHostUIPresenter _taskpaneHostUIPresenter;


        // Getting SW application reference.
        public void AssignSWApp(SldWorks inputswApp)
        {
            SolidWorksDocumentProvider._solidWorksApplication = inputswApp;
        }

        public TaskpaneHostUI()
        {
            InitializeComponent();

            _taskpaneHostUIPresenter = new TaskpaneHostUIPresenter(this);
        }

        public event EventHandler RecognizeTankAssemlby;

        public event EventHandler CreateTank;

        public event EventHandler UpdateSettings;

        public void ShowInitialConfigurationWindow(Control initialConfigControl)
        {
            Controls.Add(initialConfigControl);
            initialConfigControl.BringToFront();
        }

        public void ShowCompartmentsWindow(Control compartmentView)
        {
            try
            {
                compartmentView.Visible = false;

                Controls.Add(compartmentView);

                compartmentView.BringToFront();

                compartmentView.Visible = true;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private void RecognizeButton_Click(object sender, EventArgs e)
        {
            RecognizeTankAssemlby?.Invoke(this, e);
        }

        private void NewButton_Click(object sender, EventArgs e)
        {
            CreateTank?.Invoke(this, e);
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            UpdateSettings?.Invoke(this, e);
        }

        private void SetupButton_Click(object sender, EventArgs e)
        {
            FlangeSetup.SolidWorksSetupUtility.Run(SolidWorksDocumentProvider.GetActiveDoc());
        }
    }
}
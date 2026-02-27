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

            _taskpaneHostUIPresenter = new TaskpaneHostUIPresenter(this);
        }

        public TaskpaneHostUI(SldWorks solidWorksApp)
        {
            SolidWorksDocumentProvider._solidWorksApplication = solidWorksApp;

            InitializeComponent();

            _taskpaneHostUIPresenter = new TaskpaneHostUIPresenter(this);
        }

        public event EventHandler RecognizeTankAssemlby;

        public event EventHandler CreateTank;

        public event EventHandler UpdateSettings;

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
        
    }
}

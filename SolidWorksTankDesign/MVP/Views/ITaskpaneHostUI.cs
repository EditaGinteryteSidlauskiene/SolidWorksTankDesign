using System;
using System.Windows.Forms;

namespace MVP
{
    public interface ITaskpaneHostUI
    {
        event EventHandler RecognizeTankAssemlby;

        event EventHandler CreateTank;

        event EventHandler UpdateSettings;

        void ShowInitialConfigurationWindow(Control initialConfigControl);

        void ShowCompartmentsWindow(Control compartmentView);
    }
}

using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Helpers
{
    public static class ViewsHelper
    {
        public static void DisableControls(Control view)
        {
            foreach (Control control in view.Controls)
            {
                control.Enabled = false;
            }
        }

        public static void EnableControls(Control view)
        {
            foreach (Control control in view.Controls)
            {
                control.Enabled = true;
            }
        }
    }
}

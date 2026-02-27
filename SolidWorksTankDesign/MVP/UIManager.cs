using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.Control;

namespace SolidWorksTankDesign.MVP
{
    public static class UIManager
    {
        public static List<Control> GetAscendingControls(ControlCollection controls, string controlName)
        {
            Control[] controlsByName = controls.Find(controlName, true);
            List<Control> ascendigControls = new List<Control>();
            foreach (Control control in controlsByName)
                ascendigControls.Add(control);

            return ascendigControls
                .OrderBy(c => c.Location.Y)
                .ThenBy(c => c.Location.X)
                .ToList();
        }
    }
}

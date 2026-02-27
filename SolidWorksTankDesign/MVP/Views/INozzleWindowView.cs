using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Views
{
    public interface INozzleWindowView
    {
        void CreateCompartmentPanel();
        void AddNozzle();

        void RepositionNozzle();

        void AddNozzlePanel(Nozzle nozzle, Panel compartmentPanel);

        event EventHandler ForwardButtonPressed;

        event EventHandler BackButtonPressed;

        event EventHandler NewNozzleButtonClicked;
    }
}

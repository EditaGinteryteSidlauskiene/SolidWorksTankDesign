using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Views
{
    public interface ICompartmentWindowView
    {
        void GetInternalTreatmentsBindingSource(BindingSource internalTreatmentsBindingSource);
        void AddCompartmentPanel(CompartmentConfiguration compartmentConfiguration, int compartmentNumber);

        void ClearCompartmentPanels();

        void GetPaintingAmountUnitsDictionary(Dictionary<string, PaintingAmountUnit> paintingAmountUnitsDictionary);

        void ShowPaintingSystemsManagerView(Control paintingSystemsManagerView);

        void UpdatePaintingSystemsComboBoxes();

        void ShowNozzleWindow(Control nozzleWindow);

        event EventHandler NewCompartmentButtonClicked;

        event EventHandler PaintingSystemComboBoxSelectionChanged;

        event EventHandler ForwardButtonPressed;

        event EventHandler BackButtonPressed;
    }
}

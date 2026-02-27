using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Views
{
    public interface IPaintingSystemsManagerView
    {
        string TreatmentDescription { get; set; }
        string TreatmentCleaning { get; set; }
        string TreatmentComments { get; set; }

        object TreatmentListDataSource { get; set; }
        int SelectedTreatmentIndex { get; }

        TreatmentType TreatmentType { get; set; }

        Treatment TempTreatment { get; set; }

        event EventHandler TreatmentSelected;

        event EventHandler DeleteLayer;

        event EventHandler DeleteTreatment;

        event EventHandler SaveNewTreatment;

        event EventHandler GoBack;

        event EventHandler SaveAllChanges;

        event EventHandler CancelAllChanges;

        void SelectAndDisplayNewlyAddedTreatment();

        void DisplaySelectedTreatmentDetails(Treatment treatment);

        List<int> GetSelectedCoatingLayerIndices();

        void RefreshCoatingLayersDataGridView(Treatment treatment);

        void RefreshBidnings(ObservableCollection<Treatment> treatments);

        void ShowErrors(Dictionary<string, string> errors);
        void ShowMessage(string message, string caption = "Information", MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.Information);
    }
}

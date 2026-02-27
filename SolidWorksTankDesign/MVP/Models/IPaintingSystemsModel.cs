using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SolidWorksTankDesign.MVP.Models
{
    public interface IPaintingSystemsModel
    {
        ObservableCollection<Treatment> Treatments { get; }

        void AddTreatment(Treatment treatment);

        void DeleteTreatment(Treatment treatment);

        void ChangeComments(Treatment treatment, string comment);

        void ChangeCleaning(Treatment treatment, string cleaning);

        void ChangeDescription(Treatment treatment, string description);

        void UpdateCoatingLayer(Treatment treatment, CoatingLayer updatedLayer, int index);

        void AddCoatingLayer(Treatment treatment);

        void DeleteLayer(Treatment treatment, List<int> coatingLayersIndices);

        void SaveNewTreatment(Treatment treatment);

        bool TreatmentValidation(Treatment treatment);

        void BeginEdit();

        void EndEdit();

        void CancelEdit();
    }
}

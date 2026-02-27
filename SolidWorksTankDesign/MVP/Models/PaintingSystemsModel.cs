using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace SolidWorksTankDesign.MVP.Models
{
    public class PaintingSystemsModel : IEditableObject, IPaintingSystemsModel
    {
        public ObservableCollection<Treatment> Treatments { get; } = new ObservableCollection<Treatment>();

        private PaintingSystemsModel backupCopy;
        private bool inEdit = false;

        public PaintingSystemsModel() 
        {
            List<Treatment> allTreatments = SettingsDataManager.GetTreatments();

            Treatments.Clear();

            foreach (Treatment treatment in allTreatments)
            {
                Treatments.Add(treatment);
            }
        }

        private void SaveTreatments()
        {
            SettingsDataManager.SaveChangesInTreatments(Treatments.ToList());
        }

        public void AddTreatment(Treatment treatment)
        {
            Treatments.Add(treatment);
        }

        public void DeleteTreatment(Treatment treatment)
        {
            Treatments.Remove(treatment);

            List<Treatment> treatmentsToUpdate = new List<Treatment>();
            foreach(Treatment treatmentToSave in Treatments)
                treatmentsToUpdate.Add(treatmentToSave);

            SettingsDataManager.SaveChangesInTreatments(treatmentsToUpdate);
        }

        public void ChangeComments(Treatment treatment, string comment)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null)
            {
                targetTreatment.Comments = comment;
            }
        }

        public void ChangeCleaning(Treatment treatment, string cleaning)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null)
            {
                targetTreatment.Cleaning = cleaning;
            }
        }

        public void ChangeDescription(Treatment treatment, string description)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null)
            {
                targetTreatment.Description = description;
            }
        }

        public void UpdateCoatingLayer(Treatment treatment, CoatingLayer updatedLayer, int coatingLayerIndex)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null && coatingLayerIndex >= 0 && coatingLayerIndex < targetTreatment.CoatingLayers.Count)
            {
                targetTreatment.CoatingLayers[coatingLayerIndex] = updatedLayer;
            }
        }

        public void AddCoatingLayer(Treatment treatment)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null)
            {
                targetTreatment.CoatingLayers.Add(new CoatingLayer
                {
                    ProductName = "Product Name",
                    DryFilmThickness = 1,
                    WetFilmThickness = 1,
                    FactualConsumption = 1,
                    Units = "Units"
                });
            }
        }

        public void DeleteLayer(Treatment treatment, List<int> coatingLayersIndices)
        {
            var targetTreatment = Treatments.FirstOrDefault(t => t == treatment);
            if (targetTreatment != null)
            {
                // Sort the indices in descending order to avoid index issues after removal
                coatingLayersIndices.Sort((a, b) => b.CompareTo(a));

                foreach (int layerIndex in coatingLayersIndices)
                {
                    if (layerIndex >= 0 && layerIndex < targetTreatment.CoatingLayers.Count)
                    {
                        targetTreatment.CoatingLayers.RemoveAt(layerIndex);
                    }
                }
            }
        }

        public void SaveNewTreatment(Treatment treatment)
        {
            Treatments.Add(treatment);
        }

        public bool TreatmentValidation(Treatment treatment)
        {
            // Implement your validation logic here
            // Example:
            if (string.IsNullOrWhiteSpace(treatment.Description))
                return false;
            // Add more validation as needed

            return true;
        }

        /// <summary>
        /// Creates a deep copy of the current model for backup.
        /// </summary>
        private PaintingSystemsModel CreateDeepCopy()
        {
            var copy = new PaintingSystemsModel();

            copy.Treatments.Clear();

            foreach (var treatment in this.Treatments)
            {
                copy.Treatments.Add((Treatment)treatment.Clone());
            }

            return copy;
        }

        public void BeginEdit()
        {
            if (inEdit)
                return;

            backupCopy = this.CreateDeepCopy();
            inEdit = true;
        }

        public void EndEdit()
        {
            if (!inEdit)
                return;

            backupCopy = null;
            inEdit = false;

            SaveTreatments();
        }

        public void CancelEdit()
        {
            if (!inEdit)
                return;

            Treatments.Clear();

            foreach (var treatment in backupCopy.Treatments)
            {
                Treatments.Add((Treatment)treatment.Clone());
            }

            backupCopy = null;
            inEdit = false;
        }
    }
}

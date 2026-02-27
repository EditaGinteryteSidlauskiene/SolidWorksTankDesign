using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace SolidWorksTankDesign.Treatments
{
    public class Treatment : INotifyPropertyChanged, ICloneable
    {
        [JsonIgnore]
        private string _description;
        [JsonIgnore]
        private TreatmentType _type;
        [JsonIgnore]
        private string _cleaning;
        [JsonIgnore]
        private string _comments;

        public string Description
        {
            get => _description;
            set
            {
                _description = value;
                OnPropertyChanged(nameof(Description));
            }
        }
        public TreatmentType Type
        {
            get => _type;
            set
            {
                _type = value;
                OnPropertyChanged(nameof(Type));
            }
        }
        public string Cleaning
        {
            get => _cleaning;
            set
            {
                _cleaning = value;
                OnPropertyChanged(nameof(_cleaning));
            }
        }
        public ObservableCollection<CoatingLayer> CoatingLayers { get; set; } = new ObservableCollection<CoatingLayer>();

        public string Comments
        {
            get => _comments;
            set
            {
                _comments = value;
                OnPropertyChanged(nameof(_comments));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public object Clone()
        {
            Treatment clonedTreatment = new Treatment
            {
                Comments = this.Comments,
                Cleaning = this.Cleaning,
                Description = this.Description,
                Type = this.Type,
                CoatingLayers = new ObservableCollection<CoatingLayer>()
            };

            foreach (CoatingLayer layer in this.CoatingLayers)
            {
                clonedTreatment.CoatingLayers.Add((CoatingLayer)layer.Clone());
            }

            return clonedTreatment;
        }

        protected virtual void OnPropertyChanged(string propertyName)

        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

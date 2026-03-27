using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;

namespace SolidWorksTankDesign.TankSiteConfigurations
{
    public class CompartmentConfiguration : INotifyPropertyChanged
    {
        [JsonIgnore]
        private string _name;
        [JsonIgnore]
        private Treatment _internalSurfaceTreatment;
        [JsonIgnore]
        private double _volume;
        [JsonIgnore]
        private double _length;
        [JsonIgnore]
        private double _amount;
        [JsonIgnore]
        private PaintingAmountUnit _amountUnits;
        [JsonIgnore]
        private DishedEndAlignment _leftDishedEndAlignment;
        [JsonIgnore]
        private LeftEndConnection _leftEndConnection;

        public Guid ID { get; set; }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged(nameof(Name));
            }
        }
        public double Volume 
        { 
            get { return _volume; }
            set
            {
                _volume = Math.Round(value, 2);
                OnPropertyChanged(nameof(Volume));
            }
        }
        public double Length
        {
            get { return _length; }
            set
            {
                _length = value;
                OnPropertyChanged(nameof(Length));
            }
        }
        public Treatment InternalSurfaceTreatment 
        {
            get { return _internalSurfaceTreatment; }
            set
            {
                if (_internalSurfaceTreatment != value)
                {
                    _internalSurfaceTreatment = value;
                    OnPropertyChanged(nameof(InternalSurfaceTreatment));
                }
            }
        }
        public double Amount
        {
            get { return _amount; }
            set
            {
                _amount = value;
                OnPropertyChanged(nameof(Amount));
            }
        }
        public PaintingAmountUnit AmountUnits
        {
            get { return _amountUnits; }
            set
            {
                _amountUnits = value;
                OnPropertyChanged(nameof(AmountUnits));
            }
        }
        public DishedEndAlignment LeftDishedEndAlignment
        {
            get { return _leftDishedEndAlignment; }
            set
            {
                _leftDishedEndAlignment = value;
                OnPropertyChanged(nameof(LeftDishedEndAlignment));
            }
        }

        public ObservableCollection<NozzleConfiguration> NozzleConfigurations { get; set; } = new ObservableCollection<NozzleConfiguration>();

        public LeftEndConnection LeftEndConnection
        {
            get { return _leftEndConnection; }
            set
            {
                _leftEndConnection = value;
                OnPropertyChanged(nameof(LeftEndConnection));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public CompartmentConfiguration DeepClone()
        {
            return new CompartmentConfiguration
            {
                ID = this.ID,
                Name = this.Name,
                Volume = this.Volume,
                Length = this.Length,
                InternalSurfaceTreatment = this.InternalSurfaceTreatment,
                Amount = this.Amount,
                AmountUnits = this.AmountUnits,
                LeftDishedEndAlignment = this.LeftDishedEndAlignment,
                LeftEndConnection = this.LeftEndConnection, // Assuming LeftEndConnection is a value type or immutable

                // Deep clone the Nozzles collection
                NozzleConfigurations = new ObservableCollection<NozzleConfiguration>(
            this.NozzleConfigurations != null
            ? this.NozzleConfigurations.Select(nozzleConfiguration => nozzleConfiguration.DeepClone())
            : new List<NozzleConfiguration>())
            };
        }
    }
}

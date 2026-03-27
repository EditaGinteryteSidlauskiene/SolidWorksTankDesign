using Newtonsoft.Json;
using SolidWorksTankDesign.MVP.Enums;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SolidWorksTankDesign.TankSiteConfigurations
{
    public class NozzleConfiguration : INotifyPropertyChanged
    {
        [JsonIgnore]
        private string _designation;

        [JsonIgnore]
        public string _notes;

        [JsonIgnore]
        public NozzleReferenceType _referenceType;

        // distance along compartment axis from selected reference plane: left/right dished end or an existing nozzle in meters
        [JsonIgnore]
        public double _distanceFromReference;

        // TankCenterline / NozzleCenterline
        [JsonIgnore]
        public NozzleTopReferenceType _topReferenceType;
        [JsonIgnore]
        public NozzleBottomReferencePoint _bottomReferencePoint;

        [JsonIgnore]
        // vertical distance from the chosen top point: tank centreline OR nozzle centreline in meters
        public double _distanceFromTopReferenceMeters { get; set; } = 0.1;
        [JsonIgnore]
        public double _distanceFromBottomReferenceMeters { get; set; } = 0.1;

        // Offset and rotation
        [JsonIgnore]
        public double _offsetMeters { get; set; } = 0.1; // default 0.1 m -> 100 mm UI default
        [JsonIgnore]
        public bool _isOffsetPositive { get; set; }
        [JsonIgnore]
        public double _rotationAngleDegrees { get; set; } = 0.0;
        [JsonIgnore]
        public bool _isRotationDirectionPositive { get; set; }

        // Orientation / flip
        [JsonIgnore]
        public bool _flipped { get; set; }

        // Neck properties
        [JsonIgnore]
        public string _neckSize { get; set; }
        [JsonIgnore]  
        public double _neckThicknessMeters { get; set; }
        [JsonIgnore]  
        public string _neckMaterial { get; set; }

        // Connection properties
        [JsonIgnore]
        public string _connectionType { get; set; }
        [JsonIgnore]
        public string _connectionProperties { get; set; }
        [JsonIgnore]
        public string _connectionMaterial { get; set; }

        public Guid Id { get; set; } = Guid.NewGuid();

        public string Designation 
        {
            get { return _designation; }
            set
            {
                _designation = value;
                OnPropertyChanged(nameof(Designation));
            }
        }
        public string Notes
        {
            get { return _notes; }
            set
            {
                _notes = value;
                OnPropertyChanged(nameof(Notes));
            }
        }

        public double DistanceFromReference
        {
            get { return _distanceFromReference; }
            set
            {
                _distanceFromReference = value;
                OnPropertyChanged(nameof(DistanceFromReference));
            }
        }

        public NozzleReferenceType ReferenceType
        {
            get { return _referenceType; }
            set
            {
                _referenceType = value;
                OnPropertyChanged(nameof(ReferenceType));
            }
        }

        public NozzleTopReferenceType TopReferenceType
        {
            get { return _topReferenceType; }
            set
            {
                _topReferenceType = value;
                OnPropertyChanged(nameof(TopReferenceType));
            }
        }

        public NozzleBottomReferencePoint BottomReferencePoint
        {
            get { return _bottomReferencePoint; }
            set
            {
                _bottomReferencePoint = value;
                OnPropertyChanged(nameof(BottomReferencePoint));
            }
        }

        public double DistanceFromTopReferenceMeters
        {
            get { return _distanceFromTopReferenceMeters; }
            set
            {
                _distanceFromTopReferenceMeters = value;
                OnPropertyChanged(nameof(DistanceFromTopReferenceMeters));
            }
        }

        public double DistanceFromBottomReferenceMeters
        {
            get { return _distanceFromBottomReferenceMeters; }
            set
            {
                _distanceFromBottomReferenceMeters = value;
                OnPropertyChanged(nameof(DistanceFromBottomReferenceMeters));
            }
        }

        public double OffsetMeters
        {
            get { return _offsetMeters; }
            set
            {
                _offsetMeters = value;
                OnPropertyChanged(nameof(OffsetMeters));
            }
        }

        public bool IsOffsetPositive
        {
            get { return _isOffsetPositive; }
            set
            {
                _isOffsetPositive = value;
                OnPropertyChanged(nameof(IsOffsetPositive));
            }
        }

        public double RotationAngleDegrees
        {
            get { return _rotationAngleDegrees; }
            set
            {
                _rotationAngleDegrees = value;
                OnPropertyChanged(nameof(RotationAngleDegrees));
            }
        }

        public bool IsRotationDirectionPositive
        {
            get { return _isRotationDirectionPositive; }
            set
            {
                _isRotationDirectionPositive = value;
                OnPropertyChanged(nameof(IsRotationDirectionPositive));
            }
        }

        public bool Flipped
        {
            get { return _flipped; }
            set
            {
                _flipped = value;
                OnPropertyChanged(nameof(Flipped));
            }
        }

        public string NeckSize
        {
            get { return _neckSize; }
            set
            {
                _neckSize = value;
                OnPropertyChanged(nameof(NeckSize));
            }
        }

        public double NeckThicknessMeters
        {
            get { return _neckThicknessMeters; }
            set
            {
                _neckThicknessMeters = value;
                OnPropertyChanged(nameof(NeckThicknessMeters));
            }
        }

        public string NeckMaterial
        {
            get { return _neckMaterial; }
            set
            {
                _neckMaterial = value;
                OnPropertyChanged(nameof(NeckMaterial));
            }
        }

        public string ConnectionType
        {
            get { return _connectionType; }
            set
            {
                _connectionType = value;
                OnPropertyChanged(nameof(ConnectionType));
            }
        }
        
        public string ConnectionProperties
        {
            get { return _connectionProperties; }
            set
            {
                _connectionProperties = value;
                OnPropertyChanged(nameof(ConnectionProperties));
            }
        }

        public string ConnectionMaterial
        {
            get { return _connectionMaterial; }
            set
            {
                _connectionMaterial = value;
                OnPropertyChanged(nameof(ConnectionMaterial));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public NozzleConfiguration DeepClone()
        {
            return new NozzleConfiguration
            {
                Id = this.Id,
                Designation = this.Designation,
                Notes = this.Notes,
                ReferenceType = this.ReferenceType,
                DistanceFromReference = this.DistanceFromReference,
                TopReferenceType = this.TopReferenceType,
                BottomReferencePoint = this.BottomReferencePoint,
                DistanceFromTopReferenceMeters = this.DistanceFromTopReferenceMeters,
                DistanceFromBottomReferenceMeters = this.DistanceFromBottomReferenceMeters,
                OffsetMeters = this.OffsetMeters,
                IsOffsetPositive = this.IsOffsetPositive,
                RotationAngleDegrees = this.RotationAngleDegrees,
                IsRotationDirectionPositive = this.IsRotationDirectionPositive,
                Flipped = this.Flipped,
                NeckSize = this.NeckSize,
                NeckThicknessMeters = this.NeckThicknessMeters,
                NeckMaterial = this.NeckMaterial,
                ConnectionType = this.ConnectionType,
                ConnectionProperties = this.ConnectionProperties,
                ConnectionMaterial = this.ConnectionMaterial
                };
        }
    }
}

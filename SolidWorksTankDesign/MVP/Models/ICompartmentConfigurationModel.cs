using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Treatments;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SolidWorksTankDesign.MVP.Models
{
    public interface ICompartmentConfigurationModel
    {
        ObservableCollection <CompartmentConfiguration> CompartmentConfigurations { get;}

        ObservableCollection<Treatment> InternalTreatments { get; }

        CompartmentConfiguration CreateCompartmentConfiguration();

        void GetInternalTreatments();

        void GetCompartmentConfigurations();

        void CalculateVolume(CompartmentConfiguration compartmentConfiguration);

        void CalculateLength(CompartmentConfiguration compartmentConfiguration);

        void RemoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration);

        void UpdateLeftEndConnection(CompartmentConfiguration compartmentConfiguration, LeftEndConnection newConnection);

        void ChangeLeftDishedEndAlignment(CompartmentConfiguration compartmentConfiguration, DishedEndAlignment newAlignment);

        void MoveCompartmentConfiguration(CompartmentConfiguration compartmentConfiguration, int newIndex);

        void ModifyTank();

        void AdjustCompartmentsNames();

        void BackButtonClick();

        void ApplyCompartmentsChanges();

        bool HasChanges();
    }
}


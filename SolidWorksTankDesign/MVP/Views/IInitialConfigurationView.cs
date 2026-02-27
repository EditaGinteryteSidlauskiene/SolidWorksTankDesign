using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.Treatments;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Views
{
    public interface IInitialConfigurationView
    {
        string SerialNumber { get; set; }
        bool DefaultFolder { get; set; }
        bool SelectFolder { get; set; }
        ConstructionStandard Standard { get; set; }
        TankType TankType { get; set; }
        Class Class { get; set; }
        string Diameter { get; set; }
        string MinOperatingTemp { get; set; }
        string MaxOperatingTemp { get; set; }
        string MaxOperatingPressure { get; set; }
        string ShellLeakTestPressure { get; set; }
        Treatment ExternalSurfaceTreatment { get; set; }

        void LocateRedDotNextToSerialNumberTextBox();

        void LocateRedDotNextToDiameterTextBox();

        void LocateRedDotNextToMinOperatingTempTextBox();

        void LocateRedDotNextToMaxOperatingTempTextBox();

        void LocateRedDotNextToMaxOperatingPressureTextBox();

        void LocateRedDotNextToShellLeakTestPressureTextBox();

        void PopulateStandardsComboBox(Dictionary<string, ConstructionStandard> standardsDictionary);

        void PopulateTankTypeComboBox(Dictionary<string, TankType> tankTypesDictionary);

        void PopulateClassComboBox(Class[] classes);

        void PopulateExternalTreatmentsComboBox(List<Treatment >externalTreatments);

        void ShowCompartmentsWindow(Control compartmentWindow);

        event EventHandler TextBoxGotFocused;

        event EventHandler SaveInitialConfiguration;
    }
}

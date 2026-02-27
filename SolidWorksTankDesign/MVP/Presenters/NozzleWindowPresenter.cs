using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.TankSiteConfigurations;
using SolidWorksTankDesign.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace SolidWorksTankDesign.MVP.Presenters
{
    public class NozzleWindowPresenter
    {
        INozzleWindowView _nozzleWindowView;
        INozzleModel _nozzleModel;
        ICompartmentConfigurationModel _compartmentConfigurationModel;

        public NozzleWindowPresenter(string projectFolderPath, INozzleWindowView nozzleWindowView, INozzleModel nozzleModel, ICompartmentConfigurationModel compartmentConfigurationModel)
        {
            _nozzleWindowView = nozzleWindowView;
            _nozzleModel = nozzleModel;
            _compartmentConfigurationModel = compartmentConfigurationModel;

            _nozzleWindowView.NewNozzleButtonClicked += OnNewNozzleButtonClicked;
        }

        private void OnNewNozzleButtonClicked(object sender, EventArgs e)
        {
            System.Windows.Forms.Panel compartmentPanel = sender as System.Windows.Forms.Panel;
            CompartmentConfiguration compartmentConfiguration = (CompartmentConfiguration)compartmentPanel.Tag;

            Nozzle newNozzle = new Nozzle();
            compartmentConfiguration.Nozzles.Add(newNozzle);

            _nozzleWindowView.AddNozzlePanel(newNozzle, compartmentPanel);
        }

        public void AddNozzle(CompartmentConfigurationModel compartmentConfigurationModel, string referenceObject, string nozzleDistanceFromRef)
        {
            double.TryParse(nozzleDistanceFromRef, out double distance);

            //`_nozzleModel.AddNozzle(compartmentConfigurationModel._projectFolder, referenceObject, distance / 1000);
        }

        public void RepositionNozzle(
            bool isOffsetPositive, 
            string distanceFromFrontPlane,
            bool isRotationDirectionPositive, 
            string rotationAngle)
        {
            double distance = 0;
            double angle = 0;

            if(distanceFromFrontPlane !=  null && distanceFromFrontPlane != string.Empty)
            {
                double.TryParse(distanceFromFrontPlane, out distance);
            }

            if (rotationAngle != null && rotationAngle != string.Empty)
            {
                double.TryParse(rotationAngle, out angle);
            }

            _nozzleModel.RepositionNozzle(isOffsetPositive, distance / 1000, isRotationDirectionPositive, angle);
        }
    }
}

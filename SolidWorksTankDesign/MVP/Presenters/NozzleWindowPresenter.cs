using SolidWorksTankDesign;
using SolidWorksTankDesign.Helpers;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Views;
using SolidWorksTankDesign.MVP.Views.Controls;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Globalization;
using System.Windows.Forms;

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
            _nozzleWindowView.ForwardButtonPressed += OnForwardButtonPressed;
            _nozzleWindowView.BackButtonPressed += OnBackButtonPressed;
        }

        private void OnBackButtonPressed(object sender, EventArgs e)
        {
            // Navigate back to compartment window
            // This will be implemented based on your navigation pattern
        }

        private void OnForwardButtonPressed(object sender, EventArgs e)
        {
            // Update tank properties with all configuration changes (including nozzles)
            TankSiteDataManager.UpdateTankProperties();

            // Apply nozzle changes to the actual Nozzle objects in SolidWorks
            _nozzleModel.ApplyNozzleChanges();

            // TODO: Navigate to next window or complete the configuration
        }

        /// <summary>
        /// Execute adding all nozzles configured in the UI to their corresponding SolidWorks compartments.
        /// Iterates each CompartmentConfiguration and its Nozzles and calls the model/service for each nozzle.
        /// Continues on per-nozzle errors and shows a summary when complete.
        /// </summary>
        public void ExecuteAddNozzles()
        {
            if (_compartmentConfigurationModel == null)
            {
                MessageBox.Show("Compartment configuration model is not available.", "Operation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int total = 0;
            int succeeded = 0;
            var errors = new System.Text.StringBuilder();

            foreach (var config in _compartmentConfigurationModel.CompartmentConfigurations)
            {
                if (config.NozzleConfigurations == null) continue;

                foreach (var nozzle in config.NozzleConfigurations)
                {
                    total++;

                    // TODO: read actual reference type and distance from nozzle/UI when available.
                    // For now use sensible defaults: add relative to left dished end at 0 meters.
                    NozzleReferenceType referenceType = NozzleReferenceType.LeftDishedEnd;
                    double distanceMeters = 0.0;

                    try
                    {
                        //AddNozzle(config.ID, referenceType, distanceMeters);
                        succeeded++;
                    }
                    catch (Exception ex)
                    {
                        errors.AppendLine($"Compartment {config.Name} nozzle #{total}: {ex.Message}");
                    }
                }
            }

            string summary = $"Executed add for {total} nozzles. Successful: {succeeded}. Failed: {total - succeeded}.";
            if (errors.Length > 0)
            {
                MessageBox.Show(summary + "\n\nErrors:\n" + errors.ToString(), "Execute Add Nozzles", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                MessageBox.Show(summary, "Execute Add Nozzles", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void OnNewNozzleButtonClicked(object sender, EventArgs e)
        {
            Panel compartmentPanel = sender as Panel;
            CompartmentConfiguration compartmentConfiguration = (CompartmentConfiguration)compartmentPanel.Tag;

            NozzleConfiguration newNozzleConfig = new NozzleConfiguration();
            compartmentConfiguration.NozzleConfigurations.Add(newNozzleConfig);

            _nozzleWindowView.AddNozzlePanel(newNozzleConfig, compartmentPanel);
        }

        private void AddNozzle(
            Guid compartmentConfigId,
            NozzleReferenceType referenceType,
            string nozzleDistanceFromRef)
        {
            nozzleDistanceFromRef = nozzleDistanceFromRef.Trim();

            if (!double.TryParse(
                nozzleDistanceFromRef,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture,
                out double distance))
            {
                MessageBox.Show(
                    "Invalid input for distance. Please enter a valid number.",
                    "Input Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            else if (distance < 0)
            {
                MessageBox.Show(
                    "Distance must be a positive value.",
                    "Input Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            double distanceInMeters = distance / 1000; // Convert from millimeters to meters

            try
            {
                _nozzleModel.AddNozzle(compartmentConfigId, referenceType, distanceInMeters);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RepositionNozzle(
            bool isOffsetPositive,
            string distanceFromFrontPlane,
            bool isRotationDirectionPositive,
            string rotationAngle)
        {
            // Allow either distance or rotation (or both). Empty input means "no change" for that value.
            bool distanceProvided = !string.IsNullOrWhiteSpace(distanceFromFrontPlane);
            bool angleProvided = !string.IsNullOrWhiteSpace(rotationAngle);

            double distance = 0;
            double angle = 0;

            if (distanceProvided)
            {
                if (!double.TryParse(distanceFromFrontPlane.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out distance))
                {
                    MessageBox.Show("Invalid input for distance. Please enter a valid number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (distance < 0)
                {
                    MessageBox.Show("Distance must be a non-negative value.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (angleProvided)
            {
                if (!double.TryParse(rotationAngle.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out angle))
                {
                    MessageBox.Show("Invalid input for rotation angle. Please enter a valid number.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (angle < 0)
                {
                    MessageBox.Show("Rotation angle must be a non-negative value.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (!distanceProvided && !angleProvided)
            {
                return;
            }

            double distanceInMeters = distanceProvided ? distance / 1000.0 : 0.0; // Convert from millimeters to meters

            try
            {
                _nozzleModel.RepositionNozzle(isOffsetPositive, distanceInMeters, isRotationDirectionPositive, angle);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "Operation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        public void UpdateReferencePoint(NozzleConfiguration nozzleConfiguration, NozzleConfigurationControl.DistanceChangedEventArgs e)
        {
            if (nozzleConfiguration == null || e == null) return;

            if (e.TopReferenceType.HasValue)
            {
                nozzleConfiguration.TopReferenceType = e.TopReferenceType.Value;
            }
            else if (e.BottomReferencePoint.HasValue)
            {
                nozzleConfiguration.BottomReferencePoint = e.BottomReferencePoint.Value;
            }
        }
    }
}

using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Presenters;
using SolidWorksTankDesign.MVP.Views.Controls;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Windows.Forms;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace SolidWorksTankDesign.MVP.Views
{
    public partial class NozzleWindowView : UserControl, INozzleWindowView
    {

        private NozzleWindowPresenter _nozzleWindowPresenter;
        private CompartmentConfigurationModel _compartmentConfigurationModel;

        private readonly List<Hotspot> _hotspots = new List<Hotspot>
        {
            new Hotspot
            {
                X = 0.5589353f,
                Y = 0.1672862f,
                ReferenceType = NozzleVerticalReferenceType.TankCenterline
            },

            new Hotspot
            {
                X = 0.4372624f,
                Y = 0.1821561f,
                ReferenceType = NozzleVerticalReferenceType.NozzleCenterline
            },

            new Hotspot
            {
                X = 0.4372624f,
                Y = 0.2193309f,
                BottomReferencePoint = NozzleBottomReferencePoint.Top
            },

            new Hotspot
            {
                X = 0.4372624f,
                Y = 0.8364312f,
                BottomReferencePoint = NozzleBottomReferencePoint.Bottom
            },

            new Hotspot
            {
                X = 0.4372624f,
                Y = 0.5204461f,
                BottomReferencePoint = NozzleBottomReferencePoint.Middle
            },

            new Hotspot
            {
                X = 0.4752852f,
                Y = 0.08921933f,
                NozzlePropertiesType = NozzlePropertiesType.Neck
            },

            new Hotspot
            {
                X = 0.4372624f,
                Y = 0.03345725f,
                NozzlePropertiesType = NozzlePropertiesType.Connection
            }
        };

        private enum NozzleOrientation
        {
            Left,
            Right
        }

        private NozzleOrientation _orientation = NozzleOrientation.Left;

        private PointF _hotspot0OriginalPosition;

        private int _nozzleBottomMiddlePosition = 0;

        public NozzleWindowView(CompartmentConfigurationModel compartmentConfigurationModel)
        {
            _compartmentConfigurationModel = compartmentConfigurationModel;

            InitializeComponent();

            // Start initialization
            Initialize();

            _hotspot0OriginalPosition = new PointF(_hotspots[0].X, _hotspots[0].Y);
        }

        private void Initialize()
        {
            // Asynchronously initialize the model and presenter
            INozzleModel nozzleModel = new NozzleModel();
            _nozzleWindowPresenter = new NozzleWindowPresenter(_compartmentConfigurationModel._projectFolder, this, nozzleModel, _compartmentConfigurationModel);

            CreateCompartmentPanel();
        }

        public void CreateCompartmentPanel()
        {
            foreach (CompartmentConfiguration compartConfig in _compartmentConfigurationModel.CompartmentConfigurations)
            {
                // Add panel
                Panel newCompartmentPanel = new Panel();

                newCompartmentPanel.Tag = compartConfig;
                newCompartmentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                newCompartmentPanel.Name = "CompartmentPanel";
                newCompartmentPanel.Cursor = Cursors.Hand;
                newCompartmentPanel.Click += CompartmentPanel_Click;

                List<Control> controls = UIManager.GetAscendingControls(Controls, "CompartmentPanel");
                if (controls.Count > 0)
                {
                    newCompartmentPanel.Location = new Point(19, controls.Last().Bottom + 10);
                    newCompartmentPanel.Size = new Size(390, 25);
                }

                else
                {
                    newCompartmentPanel.Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10);
                    newCompartmentPanel.Size = new Size(390, 450);
                }

                Controls.Add(newCompartmentPanel);

                Label compartmentLabel = new Label();
                compartmentLabel.Text = compartConfig.Name;
                compartmentLabel.AutoSize = false;
                compartmentLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
                compartmentLabel.Location = new Point(4, 4);
                compartmentLabel.Name = "CompartmentLabel";
                compartmentLabel.Cursor = Cursors.Hand;
                compartmentLabel.Size = new Size(362, 20);
                compartmentLabel.Click += CompartmentLabel_Click;
                newCompartmentPanel.Controls.Add(compartmentLabel);

                Button newNozzleButton = CreateNewNozzleButton(compartConfig);
                newCompartmentPanel.Controls.Add(newNozzleButton);

                foreach (Nozzle nozzle in compartConfig.Nozzles) AddNozzlePanel(nozzle, newCompartmentPanel);
            }
        }

        private Button CreateNewNozzleButton(CompartmentConfiguration compartmentConfiguration)
        {
            Button newNozzleButton = new Button();
            newNozzleButton.Tag = compartmentConfiguration;
            newNozzleButton.AutoSize = false;
            newNozzleButton.Text = "New Nozzle";
            newNozzleButton.BackColor = System.Drawing.Color.Transparent;
            newNozzleButton.Name = "NewNozzleButton";
            newNozzleButton.Font = new Font("Microsoft Sans Serif", 9, FontStyle.Bold);
            newNozzleButton.Size = new Size(144, 29);
            newNozzleButton.Location = new Point(3, 410);
            newNozzleButton.Click += NewNozzleButton_Click;

            return newNozzleButton;
        }

        private void NewNozzleButton_Click(object sender, EventArgs e)
        {
            Control button = sender as Control;
            Control compartmentPanel = button.Parent;

            NewNozzleButtonClicked?.Invoke(compartmentPanel, EventArgs.Empty);
        }

        public void AddNozzlePanel(Nozzle nozzle, Panel compartmentPanel)
        {
            // Add panel
            Panel newNozzlePanel = new Panel();

            newNozzlePanel.Tag = nozzle;
            newNozzlePanel.BorderStyle = BorderStyle.FixedSingle;
            newNozzlePanel.Name = "NozzlePanel";
            newNozzlePanel.Cursor = Cursors.Hand;
            newNozzlePanel.Click += NozzlePanel_Click;

            List<Control> controls = UIManager.GetAscendingControls(compartmentPanel.Controls, "NozzlePanel");
            if (controls.Count > 0)
            {
                newNozzlePanel.Location = new Point(20, controls.Last().Bottom + 10);
                newNozzlePanel.Size = new Size(352, 25);
            }

            else
            {
                newNozzlePanel.Location = new Point(20, CompartmentsConfigurationLabel.Bottom + 10);
                newNozzlePanel.Size = new Size(352, 500);
            }

            compartmentPanel.Controls.Add(newNozzlePanel);

            Label nozzleLabel = new Label();
            nozzleLabel.Text = "Nozzle";
            nozzleLabel.AutoSize = false;
            nozzleLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            nozzleLabel.Location = new Point(4, 4);
            nozzleLabel.Name = "NozzleLabel";
            nozzleLabel.Cursor = Cursors.Hand;
            nozzleLabel.Size = new Size(342, 20);
            nozzleLabel.Click += NozzleLabel_Click;
            newNozzlePanel.Controls.Add(nozzleLabel);

            PictureBox nozzleSketchPictureBox = new PictureBox();
            nozzleSketchPictureBox.Name = "NozzleSketchPictureBoxDots";
            nozzleSketchPictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            nozzleSketchPictureBox.Size = new Size(250, 250);
            nozzleSketchPictureBox.Location = new Point(60, nozzleLabel.Location.Y + 200);
            nozzleSketchPictureBox.Image = Properties.Resources.NozzleSketch;
            newNozzlePanel.Controls.Add(nozzleSketchPictureBox);

            PictureBox nozzlePositionPictureBox = new PictureBox();
            nozzlePositionPictureBox.Name = "NozzlePositionPictureBoxDots";
            nozzlePositionPictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            nozzlePositionPictureBox.Size = new Size(250, 250);
            nozzlePositionPictureBox.Location = new Point(-25, 0);
            nozzlePositionPictureBox.BackColor = Color.Transparent;
            nozzlePositionPictureBox.Image = Properties.Resources.NozzlePosition;
            nozzlePositionPictureBox.MouseClick += NozzlePositionPictureBox_MouseClick;
            nozzleSketchPictureBox.Controls.Add(nozzlePositionPictureBox);

            PictureBox nozzleDistancePictureBox = new PictureBox();
            nozzleDistancePictureBox.Name = "NozzleDistancePicture";
            nozzleDistancePictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            nozzleDistancePictureBox.Size = new Size(80, 38);
            nozzleDistancePictureBox.Location = new Point(93, 90);
            nozzleDistancePictureBox.BackColor = Color.Transparent;
            nozzleDistancePictureBox.Image = Properties.Resources.DistanceFromAxis;
            nozzlePositionPictureBox.Controls.Add(nozzleDistancePictureBox);

            TextBox nozzleDistanceTextBox = new TextBox();
            nozzleDistanceTextBox.Name = "NozzleDistanceTextBox";
            nozzleDistanceTextBox.Parent = nozzleDistancePictureBox;
            nozzleDistanceTextBox.Text = "100";
            nozzleDistanceTextBox.Width = 25;
            nozzleDistanceTextBox.Height = 13;
            nozzleDistanceTextBox.BackColor = Color.White;
            nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
            nozzleDistanceTextBox.Location = new Point(20, 2);
            nozzleDistancePictureBox.Controls.Add(nozzleDistanceTextBox);

            Label flipLabel = new Label();
            flipLabel.Text = "Flip";
            flipLabel.Size = new Size(26, 18);
            flipLabel.Location = new Point(nozzleDistanceTextBox.Location.X - 2, nozzleDistanceTextBox.Location.Y +19);
            flipLabel.MouseHover += FlipLabel_MouseHover;
            flipLabel.MouseLeave += FlipLabel_MouseLeave;
            flipLabel.Click += FlipLabel_Click;
            nozzleDistancePictureBox.Controls.Add(flipLabel);

            int compartmentPanelHeight = CalculateCompartmentPanelHeight(compartmentPanel);

            compartmentPanel.Size = new Size(compartmentPanel.Size.Width, compartmentPanelHeight);
            RepositionFolowingPanels(compartmentPanel);
            RepositionLastButtons();
        }

        private void FlipLabel_Click(object sender, EventArgs e)
        {
            _orientation = _orientation == NozzleOrientation.Left 
                ? NozzleOrientation.Right
                : NozzleOrientation.Left;

            Control flipLabel = sender as Control;
            Control distancePictureBox = flipLabel.Parent;
            Control nozzlePositionPictureBox = distancePictureBox.Parent;

            if (distancePictureBox == null || nozzlePositionPictureBox == null) return;

            ApplyOrientation(nozzlePositionPictureBox, distancePictureBox);

            Control arrowToNeckProperties =
                nozzlePositionPictureBox.Controls["ArrowToNeckProperties"];
            if (arrowToNeckProperties != null) arrowToNeckProperties.BringToFront();
        }

        private void ApplyOrientation(Control nozzlePositionPictureBox, Control distancePictureBox)
        {
            UpdateMainNozzleLayout(nozzlePositionPictureBox, distancePictureBox);
            UpdateTankCentralVisualization(nozzlePositionPictureBox);
            UpdateNozzleCentralVisualization(nozzlePositionPictureBox);
            UpdateBottomVisualization(nozzlePositionPictureBox);
            UpdateHotspotsForOrientation();
        }

        private void UpdateHotspotsForOrientation()
        {
            if (_hotspots.Count == 0) return;

            if (_orientation == NozzleOrientation.Right)
            {
                // New position when flipped
                _hotspots[0].X = 0.3269962f;
                _hotspots[0].Y = 0.1710037f;
            }
            else
            {
                // Restore original position
                _hotspots[0].X = _hotspot0OriginalPosition.X;
                _hotspots[0].Y = _hotspot0OriginalPosition.Y;
            }
        }

        private void UpdateDistanceVisualization(
            PictureBox pictureBox,
            Point leftLocation,
            Size leftSize,
            Point leftChildLocation,
            Image leftImage,
            Point rightLocation,
            Size rightSize,
            Point rightChildLocation,
            Image rightImage)
        {
            if (pictureBox == null || pictureBox.Controls.Count == 0) return;

            if (_orientation == NozzleOrientation.Right)
            {
                pictureBox.Location = rightLocation;
                pictureBox.Size = rightSize;
                pictureBox.Image = rightImage;

                if (pictureBox.Controls.Count > 0)
                    pictureBox.Controls[0].Location = rightChildLocation;
            }
            else
            {
                pictureBox.Location = leftLocation;
                pictureBox.Size = leftSize;
                pictureBox.Image = leftImage;
                if (pictureBox.Controls.Count > 0)
                    pictureBox.Controls[0].Location = leftChildLocation;
            }
        }

        private void UpdateBottomVisualization(Control nozzlePositionPictureBox)
        {
            Control nozzleBottomPosition =
                nozzlePositionPictureBox.Controls["NozzleBottomPosition"];
            if (nozzleBottomPosition == null) return;

            PictureBox distanceVisualization = nozzleBottomPosition as PictureBox;

            UpdateDistanceVisualization(
                distanceVisualization,
                leftLocation: new Point(0, distanceVisualization.Location.Y),
                leftSize: new Size(100, 49),
                leftChildLocation: new Point(30, 20),
                leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
                rightLocation: new Point(123, distanceVisualization.Location.Y),
                rightSize: new Size(100, 70),
                rightChildLocation: new Point(40, 20),
                rightImage: Properties.Resources.DistanceVisualizationRightDirection2
            );
        }

        private void UpdateNozzleCentralVisualization(Control nozzlePositionPictureBox)
        {
            Control distanceFromNozzleCentralPoint =
                nozzlePositionPictureBox.Controls["DistanceFromNozzleCenterLineVisualization"];
            if (distanceFromNozzleCentralPoint == null) return;

            PictureBox distanceVisualization = distanceFromNozzleCentralPoint as PictureBox;

            UpdateDistanceVisualization(
                distanceVisualization,
                leftLocation: new Point(0, 5),
                leftSize: new Size(100, 49),
                leftChildLocation: new Point(30, 20),
                leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
                rightLocation: new Point(123, 2),
                rightSize: new Size(100, 50),
                rightChildLocation: new Point(40, 20),
                rightImage: Properties.Resources.DistanceVisualizationRightDirection2
            );
        }

        private void UpdateTankCentralVisualization(Control nozzlePositionPictureBox)
        {
            Control distanceFromTankCentralPoint =
                nozzlePositionPictureBox.Controls["DistanceFromTankCentralPointVisualization"];
            if (distanceFromTankCentralPoint == null) return;

            PictureBox distanceFromTankCentralPointPictureBox = distanceFromTankCentralPoint as PictureBox;

            UpdateDistanceVisualization(
                distanceFromTankCentralPointPictureBox,
                leftLocation: new Point(147, 0),
                leftSize: new Size(70, 75),
                leftChildLocation: new Point(37, 40),
                leftImage: Properties.Resources.DistanceVisualizationRightDirection,
                rightLocation: new Point(-20, 5),
                rightSize: new Size(100, 75),
                rightChildLocation: new Point(37, 20),
                rightImage: Properties.Resources.DistanceVisualizationLeftDirection
            );
        }

        private void UpdateMainNozzleLayout(Control nozzlePositionPictureBox, Control distancePictureBox)
        {
            if (_orientation == NozzleOrientation.Right)
            {
                nozzlePositionPictureBox.Location = new Point(31, 0);

                distancePictureBox.Location = new Point(70, 92);
                distancePictureBox.Width = 60;
                distancePictureBox.Controls[0].Location = new Point(14, distancePictureBox.Controls[0].Location.Y);
                distancePictureBox.Controls[1].Location = new Point(12, distancePictureBox.Controls[1].Location.Y);

                _hotspots[0].X = 0.3269962f;
                _hotspots[0].Y = 0.1710037f;

                Control arrowToNeckProperties =
                nozzlePositionPictureBox.Controls["ArrowToNeckProperties"];
                if (arrowToNeckProperties != null) arrowToNeckProperties.BringToFront();
            }
            else
            {
                nozzlePositionPictureBox.Location = new Point(-25, 0);

                distancePictureBox.Location = new Point(93, 90);
                distancePictureBox.Width = 80;
                distancePictureBox.Controls[0].Location = new Point(20, 2);
                distancePictureBox.Controls[1].Location = new Point(18, 21);

                _hotspots[0].X = .5589353f;
                _hotspots[0].Y = 0.1672862f;
            }
        }

        private void UpdateDistanceVisualization(
            PictureBox distancePictureBox,
            Point leftPosition, Point rightPosition,
            Size leftSize, Size rightSize,
            Image leftImage, Image rightImage,
            Point textBoxLeftPos, Point textBoxRightPos)
        {
            if (_orientation == NozzleOrientation.Left)
            {
                distancePictureBox.Location = leftPosition;
                distancePictureBox.Size = leftSize;
                distancePictureBox.Image = leftImage;
                if (distancePictureBox.Controls.Count > 0)
                    distancePictureBox.Controls[0].Location = textBoxLeftPos;
            }
            else
            {
                distancePictureBox.Location = rightPosition;
                distancePictureBox.Size = rightSize;
                distancePictureBox.Image = rightImage;
                if (distancePictureBox.Controls.Count > 0)
                    distancePictureBox.Controls[0].Location = textBoxRightPos;
            }
        }

        private void FlipLabel_MouseLeave(object sender, EventArgs e)
        {
            Label flipLabel = sender as Label;
            flipLabel.BackColor = Color.Transparent;
        }

        private void FlipLabel_MouseHover(object sender, EventArgs e)
        {
            Label flipLabel = sender as Label;
            flipLabel.BackColor = Color.LightGray;
        }

        private void NozzlePositionPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            PictureBox image = sender as PictureBox;

            var imagePoint = ControlToImagePoint(image, e.Location);

            int w = image.Image.Width;
            int h = image.Image.Height;

            float relX = imagePoint.X / w;
            float relY = imagePoint.Y / h;

            float tolerancePx = 5f; // user doesn’t need pixel-perfect clicking

            float tolX = tolerancePx / image.Image.Width;
            float tolY = tolerancePx / image.Image.Height;

            Hotspot hit = null;

            for (int i = 0; i < _hotspots.Count; i++)
            {
                float dx = relX - _hotspots[i].X;
                float dy = relY - _hotspots[i].Y;

                // normalize by tolerances so it behaves like a circle in pixels
                float nx = dx / tolX;
                float ny = dy / tolY;

                if ((nx * nx + ny * ny) <= 1f)
                {
                    hit = _hotspots[i];
                    break;
                }
            }

            if (hit != null)
            {
                if (hit.ReferenceType == NozzleVerticalReferenceType.TankCenterline)
                {
                    bool found = false;
                    foreach (Control control in image.Controls)
                    {
                        if (control.Name == "DistanceFromNozzleCenterLineVisualization")
                        {
                            control.Visible = false;
                        }

                        if (control.Name == "DistanceFromTankCentralPointVisualization")
                        {
                            found = true;
                            control.Visible = true;
                        }
                    }

                    if (found == false) AddNozzleVerticalDistanceFromTankCentralPointVisualization(image);
                }
                    
                else if (hit.ReferenceType == NozzleVerticalReferenceType.NozzleCenterline)
                {
                    bool found = false;
                    foreach (Control control in image.Controls)
                    {
                        if (control.Name == "DistanceFromTankCentralPointVisualization")
                        {
                            control.Visible = false;
                        }

                        if (control.Name == "DistanceFromNozzleCenterLineVisualization")
                        {
                            found = true;
                            control.Visible = true;
                        }
                    }

                    if (found == false) AddNozzleVerticalDistanceFromNozzleCenterLineVisualization(image);
                }

                else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Top)
                {
                    image.Image = Properties.Resources.NozzlePosition;
                    PictureBox nozzleBottomPositionPictureBox = image.Controls["NozzleBottomPosition"] as PictureBox;

                    if (nozzleBottomPositionPictureBox != null)
                    {
                        if (_orientation == NozzleOrientation.Left)
                        {
                            nozzleBottomPositionPictureBox.Location = new Point(0, 50);
                            nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationLeftDirection;
                        }
                            
                        else
                        {
                            nozzleBottomPositionPictureBox.Location = new Point(122, 45);
                            nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationRightDirection2;
                        }
                        
                    }
                    else
                    {
                        Point location = new Point(0, 50);
                        if (_orientation == NozzleOrientation.Right)
                            location = new Point(122, 50);
                        AddNozzleBottomPosition(image, location, true);
                    }
                }

                else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Bottom)
                {
                    PictureBox nozzleBottomPositionPictureBox = image.Controls["NozzleBottomPosition"] as PictureBox;

                    if (nozzleBottomPositionPictureBox != null)
                    {
                        if (_orientation == NozzleOrientation.Left)
                        {
                            PictureBox pictureBox = image as PictureBox;
                            pictureBox.Image = Properties.Resources.LongNozzlePosition;
                            nozzleBottomPositionPictureBox.Location = new Point(0, 174);
                            nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationLeftDirection;
                        }
                        else
                        {
                            PictureBox pictureBox = image as PictureBox;
                            pictureBox.Image = Properties.Resources.LongNozzlePosition;
                            nozzleBottomPositionPictureBox.Location = new Point(122, 174);
                            nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationRightDirection2;
                        }
                    }

                    else
                    {
                        Point location = new Point(0, 174);
                        if (_orientation == NozzleOrientation.Right)
                            location = new Point(122,174);
                        AddNozzleBottomPosition(image, location, false);
                    }
                }

                else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Middle)
                {
                    _nozzleBottomMiddlePosition++;

                    int locationX = 0;
                    bool isShort = false;

                    if (_orientation == NozzleOrientation.Right)
                        locationX += 122;
                    Point location;

                    if (_nozzleBottomMiddlePosition % 2 != 0)
                    {
                        image.Image = Properties.Resources.NozzlePosition;
                        location = new Point(locationX, 90);
                        isShort = true;
                    }
                    else
                    {
                        image.Image = Properties.Resources.LongNozzlePosition;
                        location = new Point(locationX, 135);
                        isShort = false;
                    }

                    Control nozzlePositionControl = image.Controls["NozzleBottomPosition"];
                    if (nozzlePositionControl != null)
                    {
                        nozzlePositionControl.Location = location;

                    } 
                    else AddNozzleBottomPosition(image, location, isShort);
                }

                else if (hit.NozzlePropertiesType == NozzlePropertiesType.Neck)
                {
                    Control sketchPicture = image.Parent;
                    Control nozzlePanel = sketchPicture.Parent;

                    Control arrowToNeckProperties = image.Controls["ArrowToConnectionProperties"];
                    if (arrowToNeckProperties != null) arrowToNeckProperties.Visible = false;

                    Control neckPropertiesPanel = nozzlePanel.Controls["ConnectionPropertiesPanel"];
                    if (neckPropertiesPanel != null) neckPropertiesPanel.Visible = false;

                    Control arrow = image.Controls["ArrowToNeckProperties"];
                    if (arrow != null && arrow.Visible == false)
                    {
                        arrow.Visible = true;
                        arrow.BringToFront();
                    }
                    

                    else if (arrow != null && arrow.Visible == true) arrow.Visible = false;

                    Control panel = nozzlePanel.Controls["NeckPropertiesPanel"];
                    if (panel != null && panel.Visible == false) panel.Visible = true;
                    else if (panel != null && panel.Visible == true) panel.Visible = false;

                    else AddNeckPropertiesPanel(image);
                }

                else if (hit.NozzlePropertiesType == NozzlePropertiesType.Connection)
                {
                    Control sketchPicture = image.Parent;
                    Control nozzlePanel = sketchPicture.Parent;

                    Control arrowToNeckProperties = image.Controls["ArrowToNeckProperties"];
                    if (arrowToNeckProperties != null) arrowToNeckProperties.Visible = false;

                    Control neckPropertiesPanel = nozzlePanel.Controls["NeckPropertiesPanel"];
                    if (neckPropertiesPanel != null) neckPropertiesPanel.Visible = false;

                    Control arrow = image.Controls["ArrowToConnectionProperties"];
                    if (arrow != null && arrow.Visible == false) arrow.Visible = true;
                    else if (arrow != null && arrow.Visible == true) arrow.Visible = false;

                    Control panel = nozzlePanel.Controls["ConnectionPropertiesPanel"];
                    if (panel != null && panel.Visible == false) panel.Visible = true;
                    else if (panel != null && panel.Visible == true) panel.Visible = false;

                    else AddConnectionPropertiesPanel(image);
                }
            }

            //MessageBox.Show($"X = {relX}, Y = {relY}");
        }

        private void AddConnectionPropertiesPanel(PictureBox pictureBox)
        {
            PictureBox arrowToPanel = new PictureBox();
            arrowToPanel.Visible = true;
            arrowToPanel.Name = "ArrowToConnectionProperties";
            arrowToPanel.SizeMode = PictureBoxSizeMode.StretchImage;
            arrowToPanel.Size = new Size(10, 10);
            arrowToPanel.Location = new Point(100, 0);
            arrowToPanel.BackColor = Color.Transparent;
            arrowToPanel.Image = Properties.Resources.ArrowToConnectionProperties;
            pictureBox.Controls.Add(arrowToPanel);


            Control sketchPicture = pictureBox.Parent;
            Control nozzlePanel = sketchPicture.Parent;

            Panel propertiesPanel = new Panel();
            propertiesPanel.Visible = true;
            propertiesPanel.Name = "ConnectionPropertiesPanel";
            propertiesPanel.Size = new Size(240, 120);
            propertiesPanel.BackColor = Color.LightGray;
            propertiesPanel.BorderStyle = BorderStyle.Fixed3D;
            propertiesPanel.Location = new Point(sketchPicture.Location.X, sketchPicture.Location.Y - 120);
            nozzlePanel.Controls.Add(propertiesPanel);

            Label ConnectionTypeLabel = new Label();
            ConnectionTypeLabel.Text = "Connection Type";
            ConnectionTypeLabel.AutoSize = false;
            ConnectionTypeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            ConnectionTypeLabel.Location = new Point(4, 4);
            ConnectionTypeLabel.Size = new Size(100, 30);
            propertiesPanel.Controls.Add(ConnectionTypeLabel);

            ComboBox connectionTypeComboBox = new ComboBox();
            connectionTypeComboBox.Name = "ConnectionTypeComboBox";
            connectionTypeComboBox.Size = new Size(100, 20);
            connectionTypeComboBox.Location = new Point(ConnectionTypeLabel.Location.X + 120, ConnectionTypeLabel.Location.Y);
            propertiesPanel.Controls.Add(connectionTypeComboBox);

            Label connectionPropertiesLabel = new Label();
            connectionPropertiesLabel.Text = "Connection Properties";
            connectionPropertiesLabel.AutoSize = false;
            connectionPropertiesLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            connectionPropertiesLabel.Location = new Point(4, 44);
            connectionPropertiesLabel.Size = new Size(100, 30);
            propertiesPanel.Controls.Add(connectionPropertiesLabel);

            ComboBox connectionPropertiesBox = new ComboBox();
            connectionPropertiesBox.Name = "ConnectionPropertiesComboBox";
            connectionPropertiesBox.Size = new Size(100, 20);
            connectionPropertiesBox.Location = new Point(connectionPropertiesLabel.Location.X + 120, connectionPropertiesLabel.Location.Y);
            propertiesPanel.Controls.Add(connectionPropertiesBox);

            Label materialLabel = new Label();
            materialLabel.Text = "Material";
            materialLabel.AutoSize = false;
            materialLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            materialLabel.Location = new Point(4, 84);
            materialLabel.Size = new Size(100, 20);
            propertiesPanel.Controls.Add(materialLabel);

            TextBox materialTextBox = new TextBox();
            materialTextBox.Name = "MaterialTextBox";
            materialTextBox.Size = new Size(100, 20);
            materialTextBox.Location = new Point(materialLabel.Location.X + 120, materialLabel.Location.Y);
            propertiesPanel.Controls.Add(materialTextBox);
        }

        private void AddNeckPropertiesPanel(PictureBox pictureBox)
        {
            PictureBox arrowToPanel = new PictureBox();
            arrowToPanel.Visible = true;
            arrowToPanel.Name = "ArrowToNeckProperties";
            arrowToPanel.SizeMode = PictureBoxSizeMode.StretchImage;
            arrowToPanel.Size = new Size(20, 20);
            arrowToPanel.Location = new Point(120, 0);
            arrowToPanel.BackColor = Color.Transparent;
            arrowToPanel.Image = Properties.Resources.ArrowToNeckProperties;
            pictureBox.Controls.Add(arrowToPanel);
            arrowToPanel.BringToFront();


            Control sketchPicture = pictureBox.Parent;
            Control nozzlePanel = sketchPicture.Parent;

            Panel propertiesPanel = new Panel();
            propertiesPanel.Visible = true;
            propertiesPanel.Name = "NeckPropertiesPanel";
            propertiesPanel.Size = new Size(240, 100);
            propertiesPanel.BackColor = Color.LightGray;
            propertiesPanel.BorderStyle = BorderStyle.Fixed3D;
            propertiesPanel.Location = new Point(sketchPicture.Location.X + 50, sketchPicture.Location.Y - 100);
            nozzlePanel.Controls.Add(propertiesPanel);

            Label sizeLabel = new Label();
            sizeLabel.Text = "Size";
            sizeLabel.AutoSize = false;
            sizeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            sizeLabel.Location = new Point(4, 4);
            sizeLabel.Size = new Size(100, 20);
            propertiesPanel.Controls.Add(sizeLabel);

            ComboBox sizeComboBox = new ComboBox();
            sizeComboBox.Name = "NeckSizeComboBox";
            sizeComboBox.Size = new Size(100, 20);
            sizeComboBox.Location = new Point(sizeLabel.Location.X + 120, sizeLabel.Location.Y);
            propertiesPanel.Controls.Add(sizeComboBox);

            Label thicknessLabel = new Label();
            thicknessLabel.Text = "Thickness";
            thicknessLabel.AutoSize = false;
            thicknessLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            thicknessLabel.Location = new Point(4, 34);
            thicknessLabel.Size = new Size(100, 20);
            propertiesPanel.Controls.Add(thicknessLabel);

            TextBox thicknessTextBox = new TextBox();
            thicknessTextBox.Name = "ThicknessTextBox";
            thicknessTextBox.Size = new Size(100, 20);
            thicknessTextBox.Location = new Point(thicknessLabel.Location.X + 120, thicknessLabel.Location.Y);
            propertiesPanel.Controls.Add(thicknessTextBox);

            Label materialLabel = new Label();
            materialLabel.Text = "Material";
            materialLabel.AutoSize = false;
            materialLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
            materialLabel.Location = new Point(4, 64);
            materialLabel.Size = new Size(100, 20);
            propertiesPanel.Controls.Add(materialLabel);

            TextBox materialTextBox = new TextBox();
            materialTextBox.Name = "MaterialTextBox";
            materialTextBox.Size = new Size(100, 20);
            materialTextBox.Location = new Point(materialLabel.Location.X + 120, materialLabel.Location.Y);
            propertiesPanel.Controls.Add(materialTextBox);
        }

        private void AddNozzleBottomPosition(PictureBox nozzlePositionPictureBox, Point location, bool isShort)
        {
            PictureBox distanceVisualization = new PictureBox();
            distanceVisualization.Name = "NozzleBottomPosition";
            nozzlePositionPictureBox.Controls.Add(distanceVisualization);

            TextBox nozzleDistanceTextBox = new TextBox();
            nozzleDistanceTextBox.Name = "NozzleDistanceFromTopToBottomTextBox";
            nozzleDistanceTextBox.Parent = distanceVisualization;
            nozzleDistanceTextBox.Text = "100";
            nozzleDistanceTextBox.Width = 30;
            nozzleDistanceTextBox.Height = 15;
            nozzleDistanceTextBox.BackColor = Color.White;
            nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
            distanceVisualization.Controls.Add(nozzleDistanceTextBox);

            UpdateDistanceVisualization(
                distancePictureBox: distanceVisualization,
                leftPosition: location,
                rightPosition: location,
                leftSize: new Size(100, 50),
                rightSize: new Size(100, 71),
                leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
                rightImage: Properties.Resources.DistanceVisualizationRightDirection2,
                textBoxLeftPos: new Point(30, 20),
                textBoxRightPos: new Point(40, 20)
            );

            // Set nozzle image based on short/long
            nozzlePositionPictureBox.Image = isShort
                ? Properties.Resources.NozzlePosition
                : Properties.Resources.LongNozzlePosition;
        }

        private void AddNozzleVerticalDistanceFromNozzleCenterLineVisualization(PictureBox nozzlePositionPictureBox)
        {
            PictureBox distanceVisualization = new PictureBox();
            distanceVisualization.Name = "DistanceFromNozzleCenterLineVisualization";
            nozzlePositionPictureBox.Controls.Add(distanceVisualization);

            TextBox nozzleDistanceTextBox = new TextBox();
            nozzleDistanceTextBox.Name = "NozzleDistanceFromNozzleCenterLineTextBox";
            nozzleDistanceTextBox.Parent = distanceVisualization;
            nozzleDistanceTextBox.Text = "100";
            nozzleDistanceTextBox.Width = 30;
            nozzleDistanceTextBox.Height = 15;
            nozzleDistanceTextBox.BackColor = Color.White;
            nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
            distanceVisualization.Controls.Add(nozzleDistanceTextBox);

            UpdateDistanceVisualization(
                distancePictureBox: distanceVisualization,
                leftPosition: new Point(0, 5),
                rightPosition: new Point(123, 2),
                leftSize: new Size(100, 49),
                rightSize: new Size(100, 50),
                leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
                rightImage: Properties.Resources.DistanceVisualizationRightDirection2,
                textBoxLeftPos: new Point(30, 20),
                textBoxRightPos: new Point(40, 20)
            );
        }

        private void AddNozzleVerticalDistanceFromTankCentralPointVisualization(PictureBox nozzlePositionPictureBox)
        {
            PictureBox distanceVisualization = new PictureBox();
            distanceVisualization.Name = "DistanceFromTankCentralPointVisualization";
            nozzlePositionPictureBox.Controls.Add(distanceVisualization);

            TextBox nozzleDistanceTextBox = new TextBox();
            nozzleDistanceTextBox.Name = "NozzleDistanceFromTankCentralPointTextBox";
            nozzleDistanceTextBox.Parent = distanceVisualization;
            nozzleDistanceTextBox.Text = "100";
            nozzleDistanceTextBox.Width = 30;
            nozzleDistanceTextBox.Height = 15;
            nozzleDistanceTextBox.BackColor = Color.White;
            nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
            distanceVisualization.Controls.Add(nozzleDistanceTextBox);

            UpdateDistanceVisualization(
                distancePictureBox: distanceVisualization,
                leftPosition: new Point(147, 3),
                rightPosition: new Point(-20, 5),
                leftSize: new Size(70, 75),
                rightSize: new Size(100, 75),
                leftImage: Properties.Resources.DistanceVisualizationRightDirection,
                rightImage: Properties.Resources.DistanceVisualizationLeftDirection,
                textBoxLeftPos: new Point(37, 17),
                textBoxRightPos: new Point(37, 20)
            );
        }

        private PointF ControlToImagePoint(PictureBox image, Point controlPoint)
        {
            var img = image.Image;

            float imageAspect = (float)img.Width / img.Height;
            float boxAspect = (float)image.ClientSize.Width / image.ClientSize.Height;

            Rectangle imageRect;

            if (boxAspect > imageAspect)
            {
                int height = image.ClientSize.Height;
                int width = (int)(height * imageAspect);
                int x = (image.ClientSize.Width - width) / 2;
                imageRect = new Rectangle(x, 0, width, height);
            }
            else
            {
                int width = image.ClientSize.Width;
                int height = (int)(width / imageAspect);
                int y = (image.ClientSize.Height - height) / 2;
                imageRect = new Rectangle(0, y, width, height);
            }

            float imageX = (controlPoint.X - imageRect.X) * img.Width / imageRect.Width;
            float imageY = (controlPoint.Y - imageRect.Y) * img.Height / imageRect.Height;

            return new PointF(imageX, imageY);
        }

        private void RepositionLastButtons()
        {
            BackArrowButton.Location = new Point(BackArrowButton.Location.X, UIManager.GetAscendingControls(Controls, "CompartmentPanel").Last().Bottom + 10);
            ForwardArrowButton.Location = new Point(ForwardArrowButton.Location.X, BackArrowButton.Location.Y);
        }

        private void RepositionNewNozzleButton(Control compartmentPanel, int buttonLocationY)
        {
            for (int i = 0; i < compartmentPanel.Controls.Count; i++)
            {
                if (compartmentPanel.Controls[i] is Button button)
                {
                    button.Location = new Point(button.Location.X, buttonLocationY);

                    return;
                }
            }
        }

        private void RepositionFolowingPanels(Control compartmentPanel)
        {
            List<Control> sortedControls = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            for (int i = 0; i < sortedControls.Count - 1; i++)
            {
                if (compartmentPanel == sortedControls[i])
                {
                    for (int j = 1; j < sortedControls.Count; j++)
                    {
                        sortedControls[j].Location = new Point(sortedControls[j-1].Location.X, sortedControls[j-1].Bottom + 10);
                    }
                }
            }
        }

        private int CalculateCompartmentPanelHeight(Panel compartmentPanel)
        {
            int compartmentPanelHeight = compartmentPanel.Height;
            int heightRequired = 0;

            foreach (Control control in compartmentPanel.Controls)
            {
                heightRequired += control.Height + 10;
            }

            if (heightRequired > compartmentPanelHeight - 30)
            {
                compartmentPanelHeight = heightRequired + 50;

                RepositionNewNozzleButton(compartmentPanel, heightRequired + 10);
            }
            return compartmentPanelHeight;
        }

        private void NozzleLabel_Click(object sender, EventArgs e)
        {
            DisplayNozzleDetails(sender);
        }

        private void NozzlePanel_Click(object sender, EventArgs e)
        {
            DisplayNozzleDetails(sender);
        }

        private void CompartmentLabel_Click(object sender, EventArgs e)
        {
            DisplayCompartmentDetails(sender);
        }

        private void CompartmentPanel_Click(object sender, EventArgs e)
        {
            DisplayCompartmentDetails(sender);
        }

        private void DisplayNozzleDetails(object controlObject)
        {
            Control control = controlObject as Control;

            if (control is Label)
            {
                control = control.Parent as Panel;
            }

            Panel compartmentPanel = control.Parent as Panel;
            List<Control> panels = UIManager.GetAscendingControls(compartmentPanel.Controls, "NozzlePanel");

            // Suspend layout for performance optimization
            this.SuspendLayout();

            try
            {
                // Step 1: Resize all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (panels[i] == control)
                    {
                        panels[i].Size = new Size(352, 500); // Enlarged panel
                        panels[i].Cursor = Cursors.Default;
                    }
                    else
                    {
                        panels[i].Size = new Size(352, 25); // Minimized panel
                        panels[i].Cursor = Cursors.Hand;
                    }
                }

                // Step 2: Reposition all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (i == 0)
                    {
                        panels[i].Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10); // First panel fixed position
                    }
                    else
                    {
                        panels[i].Location = new Point(19, panels[i - 1].Bottom + 10); // Subsequent panels
                    }
                }
            }
            finally
            {
                // Resume layout updates
                this.ResumeLayout(true);
            }
        }

        private void DisplayCompartmentDetails(object controlObject)
        {
            Control control = controlObject as Control;

            if (control is Label)
            {
                control = control.Parent as Panel;
            }

            List<Control> panels = UIManager.GetAscendingControls(Controls, "CompartmentPanel");

            // Suspend layout for performance optimization
            this.SuspendLayout();

            try
            {
                // Step 1: Resize all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (panels[i] == control)
                    {
                        int height = CalculateCompartmentPanelHeight(panels[i] as Panel);

                        panels[i].Size = new Size(390, height < 450 ? 450 : height); // Enlarged panel
                        panels[i].Cursor = Cursors.Default;

                        RepositionNewNozzleButton(panels[i], height < 450 ? 410 : height - 40);
                    }
                    else
                    {
                        panels[i].Size = new Size(390, 25); // Minimized panel
                        panels[i].Cursor = Cursors.Hand;;
                    }
                }

                // Step 2: Reposition all panels
                for (int i = 0; i < panels.Count; i++)
                {
                    if (i == 0)
                    {
                        panels[i].Location = new Point(19, CompartmentsConfigurationLabel.Bottom + 10); // First panel fixed position
                    }
                    else
                    {
                        panels[i].Location = new Point(19, panels[i - 1].Bottom + 10); // Subsequent panels
                    }
                }

                // Reposition buttons relative to the last panel
                RepositionLastButtons();
            }
            finally
            {
                // Resume layout updates
                this.ResumeLayout(true);
            }
        }

        public void AddNozzle()
        {
        //    string nozzleReference = NozzleReferenceComboBox.Text;
        //    string nozzleDistanceFromRef = NozzleReferenceDistanceTextBox.Text;

        //    _nozzleWindowPresenter.AddNozzle(_compartmentConfigurationModel, nozzleReference, nozzleDistanceFromRef);

        }

        public void RepositionNozzle()
        {
        //    string distanceFromFrontPlane = DistanceFromFrontPlaneTextBox.Text;
        //    string rotationAngle = NozzleRotationAngleTextBox.Text;

        //    _nozzleWindowPresenter.RepositionNozzle(
        //        NozzlePositiveDirectionButton.Visible, 
        //        distanceFromFrontPlane, 
        //        NozzlePositiveRotationDirectionButton.Visible,
        //        rotationAngle);
        }

        public event EventHandler ForwardButtonPressed;

        public event EventHandler BackButtonPressed;

        private void BackArrowButton_Click_1(object sender, EventArgs e)
        {
            //    //BackButtonPressed?.Invoke(this, EventArgs.Empty);
        }

        private void ForwardArrowButton_Click(object sender, EventArgs e)
        {
            //    AddNozzle();

            //    RepositionNozzle();
            //    //ForwardButtonPressed?.Invoke(this, EventArgs.Empty);

            //    //BackArrowButton.Enabled = true;
        }

        //private void NozzleNegativeDirectionButton_Click(object sender, EventArgs e)
        //{
        ////    NozzleNegativeDirectionButton.Visible = false;
        ////    NozzlePositiveDirectionButton.Visible = true;
        //}

        //private void NozzlePositiveDirectionButton_Click(object sender, EventArgs e)
        //{
        ////    NozzleNegativeDirectionButton.Visible = true;
        ////    NozzlePositiveDirectionButton.Visible = false;
        //}

        //private void NozzlePositiveRotationDirectionButton_Click(object sender, EventArgs e)
        //{
        ////    NozzleNegativeRotationDirectionButton.Visible = true;
        ////    NozzlePositiveRotationDirectionButton.Visible = false;
        //}

        //private void NozzleNegativeRotationDirectionButton_Click(object sender, EventArgs e)
        //{
        ////    NozzleNegativeRotationDirectionButton.Visible = false;
        ////    NozzlePositiveRotationDirectionButton.Visible = true;
        //}

        public event EventHandler NewNozzleButtonClicked;
    }
}

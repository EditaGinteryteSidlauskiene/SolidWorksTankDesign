using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.MVP.Models;
using SolidWorksTankDesign.MVP.Presenters;
using SolidWorksTankDesign.MVP.Services;
using SolidWorksTankDesign.MVP.Views.Controls;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using static SolidWorksTankDesign.MVP.Views.Controls.NozzleConfigurationControl;
using Point = System.Drawing.Point;
using Size = System.Drawing.Size;

namespace SolidWorksTankDesign.MVP.Views
{
    public partial class NozzleWindowView : UserControl, INozzleWindowView
    {
        private class PanelBindingContext
        {
            public NozzleConfiguration NozzleConfiguration { get; set; }
            public BindingSource BindingSource { get; set; }
        }


        private NozzleWindowPresenter _nozzleWindowPresenter;
        private CompartmentConfigurationModel _compartmentConfigurationModel;

        private readonly List<Hotspot> _hotspots = new List<Hotspot>
        {
            new Hotspot
            {
                X = 0.5589353f,
                Y = 0.1672862f,
                ReferenceType = NozzleTopReferenceType.TankCenterline
            },

            new Hotspot
            {
                X = 0.4508197f,
                Y = 0.172f,
                ReferenceType = NozzleTopReferenceType.TankCenterline
            },

            new Hotspot
            {
                X = 0.3032787f,
                Y = 0.2f,
                ReferenceType = NozzleTopReferenceType.NozzleCenterline
            },

            new Hotspot
            {
                X = 0.3032787f,
                Y = 0.24f,
                BottomReferencePoint = NozzleBottomReferencePoint.Top
            },

            new Hotspot
            {
                X = 0.3032787f,
                Y = 0.836f,
                BottomReferencePoint = NozzleBottomReferencePoint.Bottom
            },

            new Hotspot
            {
                X = 0.3032787f,
                Y = 0.532f,
                BottomReferencePoint = NozzleBottomReferencePoint.Middle
            },

            new Hotspot
            {
                X = 0.3032787f,
                Y = 0.052f,
                NozzlePropertiesType = NozzlePropertiesType.Connection
            },

            new Hotspot
            {
                X =  0.5f,
                Y = 0.5f,
                IsRotationArrow = true,
                Tolerance = 20f
            },

            new Hotspot
            {
                X = 0.45f,
                Y = 0.9f,
                IsFlipArrow = true,
                Tolerance = 25f
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
        private FlowLayoutPanel _compartmentsFlowPanel;

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
            INozzleSolidWorksService nozzleSolidWorksService = new NozzleSolidWorksService();
            INozzleModel nozzleModel = new NozzleModel(nozzleSolidWorksService);
            _nozzleWindowPresenter = new NozzleWindowPresenter(SolidWorksDocumentProvider.ProjectFolderPath, this, nozzleModel, _compartmentConfigurationModel);

            CreateCompartmentPanel();
        }

        public void CreateCompartmentPanel()
        {
            SuspendLayout();

            // Ensure single container for generated compartment panels
            if (_compartmentsFlowPanel == null)
            {
                _compartmentsFlowPanel = new FlowLayoutPanel();
                _compartmentsFlowPanel.Name = "CompartmentFlowPanel";
                _compartmentsFlowPanel.FlowDirection = FlowDirection.TopDown;
                _compartmentsFlowPanel.WrapContents = false;
                _compartmentsFlowPanel.AutoScroll = true;
                _compartmentsFlowPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

                int top = CompartmentsConfigurationLabel != null ? CompartmentsConfigurationLabel.Bottom + 10 : 30;
                _compartmentsFlowPanel.Location = new Point(19, top);
                _compartmentsFlowPanel.Size = new Size(390, this.Height - top - 60);

                Controls.Add(_compartmentsFlowPanel);
            }
            else
            {
                // Clear existing child panels
                _compartmentsFlowPanel.Controls.Clear();
            }

            System.Collections.ObjectModel.ObservableCollection<CompartmentConfiguration> configs = _compartmentConfigurationModel.CompartmentConfigurations;
            for (int idx = 0; idx < configs.Count; idx++)
            {
                CompartmentConfiguration compartConfig = configs[idx];
                // Add panel
                Panel newCompartmentPanel = new Panel();

                newCompartmentPanel.Tag = compartConfig;
                newCompartmentPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                newCompartmentPanel.Name = "CompartmentPanel";
                newCompartmentPanel.Cursor = Cursors.Hand;
                newCompartmentPanel.Click += CompartmentPanel_Click;

                // Let FlowLayoutPanel handle positioning. First compartment expanded, others collapsed.
                if (idx == 0)
                {
                    newCompartmentPanel.Size = new Size(390, 450);
                    newCompartmentPanel.Cursor = Cursors.Default;
                }
                else
                {
                    newCompartmentPanel.Size = new Size(390, 25);
                    newCompartmentPanel.Cursor = Cursors.Hand;
                }
                newCompartmentPanel.Margin = new Padding(0, 0, 0, 10);

                _compartmentsFlowPanel.Controls.Add(newCompartmentPanel);

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

                // Only add nozzle panels for the initially expanded compartment (index 0).
                if (idx == 0)
                {
                    foreach (NozzleConfiguration nozzleConfig in compartConfig.NozzleConfigurations) AddNozzlePanel(nozzleConfig, newCompartmentPanel);
                }
            }

            ResumeLayout(true);
            // Ensure navigation buttons placed after panels are created
            RepositionLastButtons();
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
            newNozzleButton.MouseClick += NewNozzleButton_MouseClick;

            return newNozzleButton;
        }

        private void NewNozzleButton_MouseClick(object sender, EventArgs e)
        {
            Control button = sender as Control;
            if (button == null) return;

            Control compartmentPanel = button.Parent as Control;
            if (compartmentPanel == null) return;

            // Raise event with the compartment panel as sender so presenter can read panel.Tag
            NewNozzleButtonClicked?.Invoke(compartmentPanel, EventArgs.Empty);
        }

        public void AddNozzlePanel(NozzleConfiguration nozzleConfiguration, Panel compartmentPanel)
        {
            PanelBindingContext panelBindingContext = new PanelBindingContext
            {
                NozzleConfiguration = nozzleConfiguration,
                BindingSource = new BindingSource { DataSource = nozzleConfiguration }
            };

            if (nozzleConfiguration == null || compartmentPanel == null) return;

            compartmentPanel.SuspendLayout();

            // Create nozzle panel
            Panel newNozzlePanel = new Panel
            {
                Tag = panelBindingContext,
                BorderStyle = BorderStyle.FixedSingle,
                Name = "NozzlePanel",
                Cursor = Cursors.Hand
            };
            newNozzlePanel.Click += NozzlePanel_Click;

            // Position: after last existing nozzle panel or after the compartment label
            List<Control> existing = UIManager.GetAscendingControls(compartmentPanel.Controls, "NozzlePanel");
            if (existing != null && existing.Count > 0)
            {
                Control last = existing.Last();
                newNozzlePanel.Location = new Point(20, last.Bottom + 10);
                newNozzlePanel.Size = new Size(352, 25);
            }
            else
            {
                // place below the compartment label if present
                Control label = compartmentPanel.Controls.Cast<Control>().FirstOrDefault(c => c.Name == "CompartmentLabel");
                int startY = label != null ? label.Bottom + 10 : 10;
                newNozzlePanel.Location = new Point(20, startY);
                newNozzlePanel.Size = new Size(352, 500);
            }

            compartmentPanel.Controls.Add(newNozzlePanel);

            // Header label for nozzle
            Label nozzleLabel = new Label
            {
                Text = "Nozzle",
                AutoSize = false,
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0))),
                Location = new Point(4, 4),
                Name = "NozzleLabel",
                Cursor = Cursors.Hand,
                Size = new Size(342, 20)
            };
            nozzleLabel.Click += NozzleLabel_Click;
            newNozzlePanel.Controls.Add(nozzleLabel);

            // Sketch image — oversized to prevent clipping when the nozzle visualization is rotated
            NozzleConfigurationControl nozzleSketchPictureBox = new NozzleConfigurationControl();
            nozzleSketchPictureBox.Name = "NozzleSketchControl";
            nozzleSketchPictureBox.Size = new Size(300, 300);
            nozzleSketchPictureBox.Location = new Point(25, nozzleLabel.Bottom + 150);

            // Add hotspots with click handlers
            foreach (Hotspot hs in _hotspots)
            {
                // create copy first, then assign OnClick to avoid capturing an unassigned local
                Hotspot hotspotCopy = new Hotspot
                {
                    X = hs.X,
                    Y = hs.Y,
                    ReferenceType = hs.ReferenceType,
                    BottomReferencePoint = hs.BottomReferencePoint,
                    NozzlePropertiesType = hs.NozzlePropertiesType,
                    IsRotationArrow = hs.IsRotationArrow,
                    IsFlipArrow = hs.IsFlipArrow,
                    Tolerance = hs.Tolerance
                };

                hotspotCopy.OnClick = () => HandleHotspotClick(hotspotCopy, nozzleSketchPictureBox, newNozzlePanel);

                nozzleSketchPictureBox.AddHotspot(hotspotCopy);
            }

            nozzleSketchPictureBox.DistanceChanged += (s, e) => _nozzleWindowPresenter.UpdateReferencePoint(nozzleConfiguration, e);

            nozzleSketchPictureBox.FlipStateChanged += (s, e) =>
            {
                nozzleConfiguration.Flipped = !nozzleConfiguration.Flipped;
                nozzleConfiguration.IsOffsetPositive = nozzleConfiguration.Flipped;
            };

            nozzleSketchPictureBox.SetNozzleConfiguration(nozzleConfiguration);

            newNozzlePanel.Controls.Add(nozzleSketchPictureBox);

            // Update container sizes and positions
            int compartmentPanelHeight = CalculateCompartmentPanelHeight(compartmentPanel);
            compartmentPanel.Size = new Size(compartmentPanel.Size.Width, compartmentPanelHeight);
            RepositionNewNozzleButton(compartmentPanel, compartmentPanelHeight < 450 ? 410 : compartmentPanelHeight - 40);
            RepositionFolowingPanels(compartmentPanel);
            RepositionLastButtons();

            compartmentPanel.ResumeLayout(true);
        }

        //private void FlipLabel_Click(object sender, EventArgs e)
        //{
        //    _orientation = _orientation == NozzleOrientation.Left 
        //        ? NozzleOrientation.Right
        //        : NozzleOrientation.Left;

        //    Control flipLabel = sender as Control;
        //    Control distancePictureBox = flipLabel.Parent;
        //    Control nozzlePositionPictureBox = distancePictureBox.Parent;

        //    if (distancePictureBox == null || nozzlePositionPictureBox == null) return;

        //    ApplyOrientation(nozzlePositionPictureBox, distancePictureBox);

        //    Control arrowToNeckProperties =
        //        nozzlePositionPictureBox.Controls["ArrowToNeckProperties"];
        //    if (arrowToNeckProperties != null) arrowToNeckProperties.BringToFront();
        //}

        //private void ApplyOrientation(Control nozzlePositionPictureBox, Control distancePictureBox)
        //{
        //    UpdateMainNozzleLayout(nozzlePositionPictureBox, distancePictureBox);
        //    UpdateTankCentralVisualization(nozzlePositionPictureBox);
        //    UpdateNozzleCentralVisualization(nozzlePositionPictureBox);
        //    UpdateBottomVisualization(nozzlePositionPictureBox);
        //    UpdateHotspotsForOrientation();
        //}

        //private void UpdateHotspotsForOrientation()
        //{
        //    if (_hotspots.Count == 0) return;

        //    if (_orientation == NozzleOrientation.Right)
        //    {
        //        // New position when flipped
        //        _hotspots[0].X = 0.3269962f;
        //        _hotspots[0].Y = 0.1710037f;
        //    }
        //    else
        //    {
        //        // Restore original position
        //        _hotspots[0].X = _hotspot0OriginalPosition.X;
        //        _hotspots[0].Y = _hotspot0OriginalPosition.Y;
        //    }
        //}

        //private void UpdateDistanceVisualization(
        //    PictureBox pictureBox,
        //    Point leftLocation,
        //    Size leftSize,
        //    Point leftChildLocation,
        //    Image leftImage,
        //    Point rightLocation,
        //    Size rightSize,
        //    Point rightChildLocation,
        //    Image rightImage)
        //{
        //    if (pictureBox == null || pictureBox.Controls.Count == 0) return;

        //    if (_orientation == NozzleOrientation.Right)
        //    {
        //        pictureBox.Location = rightLocation;
        //        pictureBox.Size = rightSize;
        //        pictureBox.Image = rightImage;

        //        if (pictureBox.Controls.Count > 0)
        //            pictureBox.Controls[0].Location = rightChildLocation;
        //    }
        //    else
        //    {
        //        pictureBox.Location = leftLocation;
        //        pictureBox.Size = leftSize;
        //        pictureBox.Image = leftImage;
        //        if (pictureBox.Controls.Count > 0)
        //            pictureBox.Controls[0].Location = leftChildLocation;
        //    }
        //}

        //private void UpdateBottomVisualization(Control nozzlePositionPictureBox)
        //{
        //    Control nozzleBottomPosition =
        //        nozzlePositionPictureBox.Controls["NozzleBottomPosition"];
        //    if (nozzleBottomPosition == null) return;

        //    PictureBox distanceVisualization = nozzleBottomPosition as PictureBox;

        //    UpdateDistanceVisualization(
        //        distanceVisualization,
        //        leftLocation: new Point(0, distanceVisualization.Location.Y),
        //        leftSize: new Size(100, 49),
        //        leftChildLocation: new Point(30, 20),
        //        leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
        //        rightLocation: new Point(123, distanceVisualization.Location.Y),
        //        rightSize: new Size(100, 70),
        //        rightChildLocation: new Point(40, 20),
        //        rightImage: Properties.Resources.DistanceVisualizationRightDirection2
        //    );
        //}

        //private void UpdateNozzleCentralVisualization(Control nozzlePositionPictureBox)
        //{
        //    Control distanceFromNozzleCentralPoint =
        //        nozzlePositionPictureBox.Controls["DistanceFromNozzleCenterLineVisualization"];
        //    if (distanceFromNozzleCentralPoint == null) return;

        //    PictureBox distanceVisualization = distanceFromNozzleCentralPoint as PictureBox;

        //    UpdateDistanceVisualization(
        //        distanceVisualization,
        //        leftLocation: new Point(0, 5),
        //        leftSize: new Size(100, 49),
        //        leftChildLocation: new Point(30, 20),
        //        leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
        //        rightLocation: new Point(123, 2),
        //        rightSize: new Size(100, 50),
        //        rightChildLocation: new Point(40, 20),
        //        rightImage: Properties.Resources.DistanceVisualizationRightDirection2
        //    );
        //}

        //private void UpdateTankCentralVisualization(Control nozzlePositionPictureBox)
        //{
        //    Control distanceFromTankCentralPoint =
        //        nozzlePositionPictureBox.Controls["DistanceFromTankCentralPointVisualization"];
        //    if (distanceFromTankCentralPoint == null) return;

        //    PictureBox distanceFromTankCentralPointPictureBox = distanceFromTankCentralPoint as PictureBox;

        //    UpdateDistanceVisualization(
        //        distanceFromTankCentralPointPictureBox,
        //        leftLocation: new Point(147, 0),
        //        leftSize: new Size(70, 75),
        //        leftChildLocation: new Point(37, 40),
        //        leftImage: Properties.Resources.DistanceVisualizationRightDirection,
        //        rightLocation: new Point(-20, 5),
        //        rightSize: new Size(100, 75),
        //        rightChildLocation: new Point(37, 20),
        //        rightImage: Properties.Resources.DistanceVisualizationLeftDirection
        //    );
        //}

        //private void UpdateMainNozzleLayout(Control nozzlePositionPictureBox, Control distancePictureBox)
        //{
        //    if (_orientation == NozzleOrientation.Right)
        //    {
        //        nozzlePositionPictureBox.Location = new Point(31, 0);

        //        distancePictureBox.Location = new Point(70, 92);
        //        distancePictureBox.Width = 60;
        //        distancePictureBox.Controls[0].Location = new Point(14, distancePictureBox.Controls[0].Location.Y);
        //        distancePictureBox.Controls[1].Location = new Point(12, distancePictureBox.Controls[1].Location.Y);

        //        _hotspots[0].X = 0.3269962f;
        //        _hotspots[0].Y = 0.1710037f;

        //        Control arrowToNeckProperties =
        //        nozzlePositionPictureBox.Controls["ArrowToNeckProperties"];
        //        if (arrowToNeckProperties != null) arrowToNeckProperties.BringToFront();
        //    }
        //    else
        //    {
        //        nozzlePositionPictureBox.Location = new Point(-25, 0);

        //        distancePictureBox.Location = new Point(93, 90);
        //        distancePictureBox.Width = 80;
        //        distancePictureBox.Controls[0].Location = new Point(20, 2);
        //        distancePictureBox.Controls[1].Location = new Point(18, 21);

        //        _hotspots[0].X = .5589353f;
        //        _hotspots[0].Y = 0.1672862f;
        //    }
        //}

        //private void UpdateDistanceVisualization(
        //    PictureBox distancePictureBox,
        //    Point leftPosition, Point rightPosition,
        //    Size leftSize, Size rightSize,
        //    Image leftImage, Image rightImage,
        //    Point textBoxLeftPos, Point textBoxRightPos)
        //{
        //    if (_orientation == NozzleOrientation.Left)
        //    {
        //        distancePictureBox.Location = leftPosition;
        //        distancePictureBox.Size = leftSize;
        //        distancePictureBox.Image = leftImage;
        //        if (distancePictureBox.Controls.Count > 0)
        //            distancePictureBox.Controls[0].Location = textBoxLeftPos;
        //    }
        //    else
        //    {
        //        distancePictureBox.Location = rightPosition;
        //        distancePictureBox.Size = rightSize;
        //        distancePictureBox.Image = rightImage;
        //        if (distancePictureBox.Controls.Count > 0)
        //            distancePictureBox.Controls[0].Location = textBoxRightPos;
        //    }
        //}

        //private void FlipLabel_MouseLeave(object sender, EventArgs e)
        //{
        //    Label flipLabel = sender as Label;
        //    flipLabel.BackColor = Color.Transparent;
        //}

        //private void FlipLabel_MouseHover(object sender, EventArgs e)
        //{
        //    Label flipLabel = sender as Label;
        //    flipLabel.BackColor = Color.LightGray;
        //}

        //private void NozzlePositionPictureBox_MouseClick(object sender, MouseEventArgs e)
        //{
        //    PictureBox image = sender as PictureBox;

        //    var imagePoint = ControlToImagePoint(image, e.Location);

        //    int w = image.Image.Width;
        //    int h = image.Image.Height;

        //    float relX = imagePoint.X / w;
        //    float relY = imagePoint.Y / h;

        //    float tolerancePx = 5f; // user doesn’t need pixel-perfect clicking

        //    float tolX = tolerancePx / image.Image.Width;
        //    float tolY = tolerancePx / image.Image.Height;

        //    Hotspot hit = null;

        //    for (int i = 0; i < _hotspots.Count; i++)
        //    {
        //        float dx = relX - _hotspots[i].X;
        //        float dy = relY - _hotspots[i].Y;

        //        // normalize by tolerances so it behaves like a circle in pixels
        //        float nx = dx / tolX;
        //        float ny = dy / tolY;

        //        if ((nx * nx + ny * ny) <= 1f)
        //        {
        //            hit = _hotspots[i];
        //            break;
        //        }
        //    }

        //    if (hit != null)
        //    {
        //        if (hit.ReferenceType == NozzleTopReferenceType.TankCenterline)
        //        {
        //            bool found = false;
        //            foreach (Control control in image.Controls)
        //            {
        //                if (control.Name == "DistanceFromNozzleCenterLineVisualization")
        //                {
        //                    control.Visible = false;
        //                }

        //                if (control.Name == "DistanceFromTankCentralPointVisualization")
        //                {
        //                    found = true;
        //                    control.Visible = true;
        //                }
        //            }

        //            if (found == false) AddNozzleVerticalDistanceFromTankCentralPointVisualization(image);
        //        }
                    
        //        else if (hit.ReferenceType == NozzleTopReferenceType.NozzleCenterline)
        //        {
        //            bool found = false;
        //            foreach (Control control in image.Controls)
        //            {
        //                if (control.Name == "DistanceFromTankCentralPointVisualization")
        //                {
        //                    control.Visible = false;
        //                }

        //                if (control.Name == "DistanceFromNozzleCenterLineVisualization")
        //                {
        //                    found = true;
        //                    control.Visible = true;
        //                }
        //            }

        //            if (found == false) AddNozzleVerticalDistanceFromNozzleCenterLineVisualization(image);
        //        }

        //        else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Top)
        //        {
        //            image.Image = Properties.Resources.NozzlePosition;
        //            PictureBox nozzleBottomPositionPictureBox = image.Controls["NozzleBottomPosition"] as PictureBox;

        //            if (nozzleBottomPositionPictureBox != null)
        //            {
        //                if (_orientation == NozzleOrientation.Left)
        //                {
        //                    nozzleBottomPositionPictureBox.Location = new Point(0, 50);
        //                    nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationLeftDirection;
        //                }
                            
        //                else
        //                {
        //                    nozzleBottomPositionPictureBox.Location = new Point(122, 45);
        //                    nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationRightDirection2;
        //                }
                        
        //            }
        //            else
        //            {
        //                Point location = new Point(0, 50);
        //                if (_orientation == NozzleOrientation.Right)
        //                    location = new Point(122, 50);
        //                AddNozzleBottomPosition(image, location, true);
        //            }
        //        }

        //        else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Bottom)
        //        {
        //            PictureBox nozzleBottomPositionPictureBox = image.Controls["NozzleBottomPosition"] as PictureBox;

        //            if (nozzleBottomPositionPictureBox != null)
        //            {
        //                if (_orientation == NozzleOrientation.Left)
        //                {
        //                    PictureBox pictureBox = image as PictureBox;
        //                    pictureBox.Image = Properties.Resources.LongNozzlePosition;
        //                    nozzleBottomPositionPictureBox.Location = new Point(0, 174);
        //                    nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationLeftDirection;
        //                }
        //                else
        //                {
        //                    PictureBox pictureBox = image as PictureBox;
        //                    pictureBox.Image = Properties.Resources.LongNozzlePosition;
        //                    nozzleBottomPositionPictureBox.Location = new Point(122, 174);
        //                    nozzleBottomPositionPictureBox.Image = Properties.Resources.DistanceVisualizationRightDirection2;
        //                }
        //            }

        //            else
        //            {
        //                Point location = new Point(0, 174);
        //                if (_orientation == NozzleOrientation.Right)
        //                    location = new Point(122,174);
        //                AddNozzleBottomPosition(image, location, false);
        //            }
        //        }

        //        else if (hit.BottomReferencePoint == NozzleBottomReferencePoint.Middle)
        //        {
        //            _nozzleBottomMiddlePosition++;

        //            int locationX = 0;
        //            bool isShort = false;

        //            if (_orientation == NozzleOrientation.Right)
        //                locationX += 122;
        //            Point location;

        //            if (_nozzleBottomMiddlePosition % 2 != 0)
        //            {
        //                image.Image = Properties.Resources.NozzlePosition;
        //                location = new Point(locationX, 90);
        //                isShort = true;
        //            }
        //            else
        //            {
        //                image.Image = Properties.Resources.LongNozzlePosition;
        //                location = new Point(locationX, 135);
        //                isShort = false;
        //            }

        //            Control nozzlePositionControl = image.Controls["NozzleBottomPosition"];
        //            if (nozzlePositionControl != null)
        //            {
        //                nozzlePositionControl.Location = location;

        //            } 
        //            else AddNozzleBottomPosition(image, location, isShort);
        //        }

        //        else if (hit.NozzlePropertiesType == NozzlePropertiesType.Neck)
        //        {
        //            Control sketchPicture = image.Parent;
        //            Control nozzlePanel = sketchPicture.Parent;

        //            Control arrowToNeckProperties = image.Controls["ArrowToConnectionProperties"];
        //            if (arrowToNeckProperties != null) arrowToNeckProperties.Visible = false;

        //            Control neckPropertiesPanel = nozzlePanel.Controls["ConnectionPropertiesPanel"];
        //            if (neckPropertiesPanel != null) neckPropertiesPanel.Visible = false;

        //            Control arrow = image.Controls["ArrowToNeckProperties"];
        //            if (arrow != null && arrow.Visible == false)
        //            {
        //                arrow.Visible = true;
        //                arrow.BringToFront();
        //            }
                    

        //            else if (arrow != null && arrow.Visible == true) arrow.Visible = false;

        //            Control panel = nozzlePanel.Controls["NeckPropertiesPanel"];
        //            if (panel != null && panel.Visible == false) panel.Visible = true;
        //            else if (panel != null && panel.Visible == true) panel.Visible = false;

        //            else AddNeckPropertiesPanel(image);
        //        }

        //        else if (hit.NozzlePropertiesType == NozzlePropertiesType.Connection)
        //        {
        //            Control sketchPicture = image.Parent;
        //            Control nozzlePanel = sketchPicture.Parent;

        //            Control arrowToNeckProperties = image.Controls["ArrowToNeckProperties"];
        //            if (arrowToNeckProperties != null) arrowToNeckProperties.Visible = false;

        //            Control neckPropertiesPanel = nozzlePanel.Controls["NeckPropertiesPanel"];
        //            if (neckPropertiesPanel != null) neckPropertiesPanel.Visible = false;

        //            Control arrow = image.Controls["ArrowToConnectionProperties"];
        //            if (arrow != null && arrow.Visible == false) arrow.Visible = true;
        //            else if (arrow != null && arrow.Visible == true) arrow.Visible = false;

        //            Control panel = nozzlePanel.Controls["ConnectionPropertiesPanel"];
        //            if (panel != null && panel.Visible == false) panel.Visible = true;
        //            else if (panel != null && panel.Visible == true) panel.Visible = false;

        //            else AddConnectionPropertiesPanel(image);
        //        }
        //    }

        //    //MessageBox.Show($"X = {relX}, Y = {relY}");
        //}

        //private void AddConnectionPropertiesPanel(PictureBox pictureBox)
        //{
        //    PictureBox arrowToPanel = new PictureBox();
        //    arrowToPanel.Visible = true;
        //    arrowToPanel.Name = "ArrowToConnectionProperties";
        //    arrowToPanel.SizeMode = PictureBoxSizeMode.StretchImage;
        //    arrowToPanel.Size = new Size(10, 10);
        //    arrowToPanel.Location = new Point(100, 0);
        //    arrowToPanel.BackColor = Color.Transparent;
        //    arrowToPanel.Image = Properties.Resources.ArrowToConnectionProperties;
        //    pictureBox.Controls.Add(arrowToPanel);


        //    Control sketchPicture = pictureBox.Parent;
        //    Control nozzlePanel = sketchPicture.Parent;

        //    Panel propertiesPanel = new Panel();
        //    propertiesPanel.Visible = true;
        //    propertiesPanel.Name = "ConnectionPropertiesPanel";
        //    propertiesPanel.Size = new Size(240, 120);
        //    propertiesPanel.BackColor = Color.LightGray;
        //    propertiesPanel.BorderStyle = BorderStyle.Fixed3D;
        //    propertiesPanel.Location = new Point(sketchPicture.Location.X, sketchPicture.Location.Y - 120);
        //    nozzlePanel.Controls.Add(propertiesPanel);

        //    Label ConnectionTypeLabel = new Label();
        //    ConnectionTypeLabel.Text = "Connection Type";
        //    ConnectionTypeLabel.AutoSize = false;
        //    ConnectionTypeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    ConnectionTypeLabel.Location = new Point(4, 4);
        //    ConnectionTypeLabel.Size = new Size(100, 30);
        //    propertiesPanel.Controls.Add(ConnectionTypeLabel);

        //    ComboBox connectionTypeComboBox = new ComboBox();
        //    connectionTypeComboBox.Name = "ConnectionTypeComboBox";
        //    connectionTypeComboBox.Size = new Size(100, 20);
        //    connectionTypeComboBox.Location = new Point(ConnectionTypeLabel.Location.X + 120, ConnectionTypeLabel.Location.Y);
        //    propertiesPanel.Controls.Add(connectionTypeComboBox);

        //    Label connectionPropertiesLabel = new Label();
        //    connectionPropertiesLabel.Text = "Connection Properties";
        //    connectionPropertiesLabel.AutoSize = false;
        //    connectionPropertiesLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    connectionPropertiesLabel.Location = new Point(4, 44);
        //    connectionPropertiesLabel.Size = new Size(100, 30);
        //    propertiesPanel.Controls.Add(connectionPropertiesLabel);

        //    ComboBox connectionPropertiesBox = new ComboBox();
        //    connectionPropertiesBox.Name = "ConnectionPropertiesComboBox";
        //    connectionPropertiesBox.Size = new Size(100, 20);
        //    connectionPropertiesBox.Location = new Point(connectionPropertiesLabel.Location.X + 120, connectionPropertiesLabel.Location.Y);
        //    propertiesPanel.Controls.Add(connectionPropertiesBox);

        //    Label connectionMaterialLabel = new Label();
        //    connectionMaterialLabel.Text = "Material";
        //    connectionMaterialLabel.AutoSize = false;
        //    connectionMaterialLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    connectionMaterialLabel.Location = new Point(4, 84);
        //    connectionMaterialLabel.Size = new Size(100, 20);
        //    propertiesPanel.Controls.Add(connectionMaterialLabel);

        //    TextBox materialTextBox = new TextBox();
        //    materialTextBox.Name = "MaterialTextBox";
        //    materialTextBox.Size = new Size(100, 20);
        //    materialTextBox.Location = new Point(connectionMaterialLabel.Location.X + 120, connectionMaterialLabel.Location.Y);
        //    propertiesPanel.Controls.Add(materialTextBox);
        //}

        //private void AddNeckPropertiesPanel(PictureBox pictureBox)
        //{
        //    PictureBox arrowToPanel = new PictureBox();
        //    arrowToPanel.Visible = true;
        //    arrowToPanel.Name = "ArrowToNeckProperties";
        //    arrowToPanel.SizeMode = PictureBoxSizeMode.StretchImage;
        //    arrowToPanel.Size = new Size(20, 20);
        //    arrowToPanel.Location = new Point(120, 0);
        //    arrowToPanel.BackColor = Color.Transparent;
        //    arrowToPanel.Image = Properties.Resources.ArrowToNeckProperties;
        //    pictureBox.Controls.Add(arrowToPanel);
        //    arrowToPanel.BringToFront();


        //    Control sketchPicture = pictureBox.Parent;
        //    Control nozzlePanel = sketchPicture.Parent;

        //    Panel propertiesPanel = new Panel();
        //    propertiesPanel.Visible = true;
        //    propertiesPanel.Name = "NeckPropertiesPanel";
        //    propertiesPanel.Size = new Size(240, 100);
        //    propertiesPanel.BackColor = Color.LightGray;
        //    propertiesPanel.BorderStyle = BorderStyle.Fixed3D;
        //    propertiesPanel.Location = new Point(sketchPicture.Location.X + 50, sketchPicture.Location.Y - 100);
        //    nozzlePanel.Controls.Add(propertiesPanel);

        //    Label sizeLabel = new Label();
        //    sizeLabel.Text = "Size";
        //    sizeLabel.AutoSize = false;
        //    sizeLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    sizeLabel.Location = new Point(4, 4);
        //    sizeLabel.Size = new Size(100, 20);
        //    propertiesPanel.Controls.Add(sizeLabel);

        //    ComboBox sizeComboBox = new ComboBox();
        //    sizeComboBox.Name = "NeckSizeComboBox";
        //    sizeComboBox.Size = new Size(100, 20);
        //    sizeComboBox.Location = new Point(sizeLabel.Location.X + 120, sizeLabel.Location.Y);
        //    propertiesPanel.Controls.Add(sizeComboBox);

        //    Label thicknessLabel = new Label();
        //    thicknessLabel.Text = "Thickness";
        //    thicknessLabel.AutoSize = false;
        //    thicknessLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    thicknessLabel.Location = new Point(4, 34);
        //    thicknessLabel.Size = new Size(100, 20);
        //    propertiesPanel.Controls.Add(thicknessLabel);

        //    TextBox thicknessTextBox = new TextBox();
        //    thicknessTextBox.Name = "ThicknessTextBox";
        //    thicknessTextBox.Size = new Size(100, 20);
        //    thicknessTextBox.Location = new Point(thicknessLabel.Location.X + 120, thicknessLabel.Location.Y);
        //    propertiesPanel.Controls.Add(thicknessTextBox);

        //    Label connectionMaterialLabel = new Label();
        //    connectionMaterialLabel.Text = "Material";
        //    connectionMaterialLabel.AutoSize = false;
        //    connectionMaterialLabel.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, ((byte)(0)));
        //    connectionMaterialLabel.Location = new Point(4, 64);
        //    connectionMaterialLabel.Size = new Size(100, 20);
        //    propertiesPanel.Controls.Add(connectionMaterialLabel);

        //    TextBox materialTextBox = new TextBox();
        //    materialTextBox.Name = "MaterialTextBox";
        //    materialTextBox.Size = new Size(100, 20);
        //    materialTextBox.Location = new Point(connectionMaterialLabel.Location.X + 120, connectionMaterialLabel.Location.Y);
        //    propertiesPanel.Controls.Add(materialTextBox);
        //}

        //private void AddNozzleBottomPosition(PictureBox nozzlePositionPictureBox, Point location, bool isShort)
        //{
        //    PictureBox distanceVisualization = new PictureBox();
        //    distanceVisualization.Name = "NozzleBottomPosition";
        //    nozzlePositionPictureBox.Controls.Add(distanceVisualization);

        //    TextBox nozzleDistanceTextBox = new TextBox();
        //    nozzleDistanceTextBox.Name = "NozzleDistanceFromTopToBottomTextBox";
        //    nozzleDistanceTextBox.Parent = distanceVisualization;
        //    nozzleDistanceTextBox.Text = "100";
        //    nozzleDistanceTextBox.Width = 30;
        //    nozzleDistanceTextBox.Height = 15;
        //    nozzleDistanceTextBox.BackColor = Color.White;
        //    nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
        //    distanceVisualization.Controls.Add(nozzleDistanceTextBox);

        //    UpdateDistanceVisualization(
        //        distancePictureBox: distanceVisualization,
        //        leftPosition: location,
        //        rightPosition: location,
        //        leftSize: new Size(100, 50),
        //        rightSize: new Size(100, 71),
        //        leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
        //        rightImage: Properties.Resources.DistanceVisualizationRightDirection2,
        //        textBoxLeftPos: new Point(30, 20),
        //        textBoxRightPos: new Point(40, 20)
        //    );

        //    // Set nozzle image based on short/long
        //    nozzlePositionPictureBox.Image = isShort
        //        ? Properties.Resources.NozzlePosition
        //        : Properties.Resources.LongNozzlePosition;
        //}

        //private void AddNozzleVerticalDistanceFromNozzleCenterLineVisualization(PictureBox nozzlePositionPictureBox)
        //{
        //    PictureBox distanceVisualization = new PictureBox();
        //    distanceVisualization.Name = "DistanceFromNozzleCenterLineVisualization";
        //    nozzlePositionPictureBox.Controls.Add(distanceVisualization);

        //    TextBox nozzleDistanceTextBox = new TextBox();
        //    nozzleDistanceTextBox.Name = "NozzleDistanceFromNozzleCenterLineTextBox";
        //    nozzleDistanceTextBox.Parent = distanceVisualization;
        //    nozzleDistanceTextBox.Text = "100";
        //    nozzleDistanceTextBox.Width = 30;
        //    nozzleDistanceTextBox.Height = 15;
        //    nozzleDistanceTextBox.BackColor = Color.White;
        //    nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
        //    distanceVisualization.Controls.Add(nozzleDistanceTextBox);

        //    UpdateDistanceVisualization(
        //        distancePictureBox: distanceVisualization,
        //        leftPosition: new Point(0, 5),
        //        rightPosition: new Point(123, 2),
        //        leftSize: new Size(100, 49),
        //        rightSize: new Size(100, 50),
        //        leftImage: Properties.Resources.DistanceVisualizationLeftDirection,
        //        rightImage: Properties.Resources.DistanceVisualizationRightDirection2,
        //        textBoxLeftPos: new Point(30, 20),
        //        textBoxRightPos: new Point(40, 20)
        //    );
        //}

        //private void AddNozzleVerticalDistanceFromTankCentralPointVisualization(PictureBox nozzlePositionPictureBox)
        //{
        //    PictureBox distanceVisualization = new PictureBox();
        //    distanceVisualization.Name = "DistanceFromTankCentralPointVisualization";
        //    nozzlePositionPictureBox.Controls.Add(distanceVisualization);

        //    TextBox nozzleDistanceTextBox = new TextBox();
        //    nozzleDistanceTextBox.Name = "NozzleDistanceFromTankCentralPointTextBox";
        //    nozzleDistanceTextBox.Parent = distanceVisualization;
        //    nozzleDistanceTextBox.Text = "100";
        //    nozzleDistanceTextBox.Width = 30;
        //    nozzleDistanceTextBox.Height = 15;
        //    nozzleDistanceTextBox.BackColor = Color.White;
        //    nozzleDistanceTextBox.BorderStyle = BorderStyle.None;
        //    distanceVisualization.Controls.Add(nozzleDistanceTextBox);

        //    UpdateDistanceVisualization(
        //        distancePictureBox: distanceVisualization,
        //        leftPosition: new Point(147, 3),
        //        rightPosition: new Point(-20, 5),
        //        leftSize: new Size(70, 75),
        //        rightSize: new Size(100, 75),
        //        leftImage: Properties.Resources.DistanceVisualizationRightDirection,
        //        rightImage: Properties.Resources.DistanceVisualizationLeftDirection,
        //        textBoxLeftPos: new Point(37, 17),
        //        textBoxRightPos: new Point(37, 20)
        //    );
        //}

        //private PointF ControlToImagePoint(PictureBox image, Point controlPoint)
        //{
        //    var img = image.Image;

        //    float imageAspect = (float)img.Width / img.Height;
        //    float boxAspect = (float)image.ClientSize.Width / image.ClientSize.Height;

        //    Rectangle imageRect;

        //    if (boxAspect > imageAspect)
        //    {
        //        int height = image.ClientSize.Height;
        //        int width = (int)(height * imageAspect);
        //        int x = (image.ClientSize.Width - width) / 2;
        //        imageRect = new Rectangle(x, 0, width, height);
        //    }
        //    else
        //    {
        //        int width = image.ClientSize.Width;
        //        int height = (int)(width / imageAspect);
        //        int y = (image.ClientSize.Height - height) / 2;
        //        imageRect = new Rectangle(0, y, width, height);
        //    }

        //    float imageX = (controlPoint.X - imageRect.X) * img.Width / imageRect.Width;
        //    float imageY = (controlPoint.Y - imageRect.Y) * img.Height / imageRect.Height;

        //    return new PointF(imageX, imageY);
        //}

        private void RepositionLastButtons()
        {
            // Determine container to search for compartment panels
            Control container = _compartmentsFlowPanel ?? (Control)this;

            var panels = UIManager.GetAscendingControls(container.Controls, "CompartmentPanel");
            if (panels == null || panels.Count == 0) return;

            // Move buttons further down: use +50 from last panel bottom
            BackArrowButton.Location = new Point(BackArrowButton.Location.X, panels.Last().Bottom + 50);
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
                        var panel = panels[i] as Panel;

                        // If this panel doesn't yet contain nozzle panels, add them now from model
                        var config = panel.Tag as CompartmentConfiguration;
                        bool hasNozzlePanels = panel.Controls.Cast<Control>().Any(c => c.Name == "NozzlePanel");
                        if (!hasNozzlePanels && config != null)
                        {
                            foreach (var nozzleConfig in config.NozzleConfigurations)
                            {
                                AddNozzlePanel(nozzleConfig, panel);
                            }
                        }

                        int height = CalculateCompartmentPanelHeight(panel);

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

            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0].ActivateDocument();
            Feature refPlane = SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0].GetRightEndPlane();


            SolidWorksDocumentProvider._tankSiteAssembly._compartmentsManager.Compartments[0].AddNozzle(
                "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Nozzle position sketch.SLDASM",
                "C:\\Users\\Edita\\TankDesignStudio\\TankSite\\Manholes\\Manhole DN600 Neck with flange.SLDASM",
                0,
                refPlane,
                1,
                true,
                2500);




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

        private void HandleHotspotClick(Hotspot hotspot, NozzleConfigurationControl sketch, Panel nozzlePanel)
        {
            if (hotspot.ReferenceType == NozzleTopReferenceType.TankCenterline)
            {
                sketch.ToggleTankCenterlineDistanceVisualization();
            }
            else if (hotspot.ReferenceType == NozzleTopReferenceType.NozzleCenterline)
            {
                sketch.ToggleNozzleCenterlineDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Top)
            {
                sketch.ToggleNozzleTopDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Bottom)
            {
                sketch.ToggleNozzleBottomDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Middle)
            {
                sketch.ToggleNozzleMiddleDistanceVisualization();
            }
            else if (hotspot.NozzlePropertiesType == NozzlePropertiesType.Connection)
            {
                ToggleConnectionPropertiesPanel(sketch, nozzlePanel);
            }
            else if (hotspot.IsRotationArrow == true)
            {
                sketch.ToggleRotationVisualization();
            }
            else if (hotspot.IsFlipArrow == true)
            {
                sketch.ToggleFlip();
            }
        }

                private void ToggleConnectionPropertiesPanel(NozzleConfigurationControl sketch, Panel nozzlePanel)
        {
            PanelBindingContext nozzlePanelBindingContext = nozzlePanel.Tag as PanelBindingContext;
            BindingSource bindingSource = nozzlePanelBindingContext.BindingSource;
            if (bindingSource == null) return;

            // Hide neck properties panel if visible
            Control neckPanel = nozzlePanel.Controls.Cast<Control>()
                .FirstOrDefault(c => c.Name == "NeckPropertiesPanel");
            if (neckPanel != null)
            {
                neckPanel.Visible = false;
            }

            // Check if connection panel already exists
            Control existingPanel = nozzlePanel.Controls.Cast<Control>()
                .FirstOrDefault(c => c.Name == "ConnectionPropertiesPanel");

            if (existingPanel != null)
            {
                // Toggle visibility
                existingPanel.Visible = !existingPanel.Visible;
                sketch.SetVisualizationsEnabled(!existingPanel.Visible);
                return;
            }

            // Create new connection properties panel
            Panel propertiesPanel = new Panel
            {
                Name = "ConnectionPropertiesPanel",
                Size = new Size(250, 220),  
                BackColor = Color.LightGray,
                BorderStyle = BorderStyle.FixedSingle,
                Location = new Point(
                    (sketch.Width - 240) / 2,
                    (sketch.Height - 180) / 2)
            };

            // Add close button at top-right corner
            Button closeButton = new Button
            {
                Text = "X",
                Size = new Size(20, 20),
                Location = new Point(250 - 22, 2),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
                Font = new Font("Arial", 8, FontStyle.Bold)
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.FlatAppearance.MouseOverBackColor = Color.Silver;
            closeButton.Click += (s, ev) =>
            {
                propertiesPanel.Visible = false;
                sketch.SetVisualizationsEnabled(true);
            };
            propertiesPanel.Controls.Add(closeButton);

            // Add "Size" label and combobox
            Label sizeLabel = new Label
            {
                Text = "Size",
                Location = new Point(4, 6),
                Size = new Size(100, 20),
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(sizeLabel);

            ComboBox sizeComboBox = new ComboBox
            {
                Name = "NeckSizeComboBox",
                Size = new Size(100, 20),
                Location = new Point(124, 6),
                DataBindings = { new Binding("Text", bindingSource, "NeckSize", true, DataSourceUpdateMode.OnPropertyChanged) }
            };
            sizeComboBox.Items.AddRange(new object[] { "Size one", "Size two", "Size three" });
            propertiesPanel.Controls.Add(sizeComboBox);

            // Add "Connection Type" label and combobox
            Label connectionTypeLabel = new Label
            {
                Text = "Connection Type",
                AutoSize = false,
                Location = new Point(4, sizeLabel.Bottom + 10),
                Size = new Size(100, 30),
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(connectionTypeLabel);

            ComboBox connectionTypeComboBox = new ComboBox
            {
                Name = "ConnectionTypeComboBox",
                Size = new Size(100, 20),
                Location = new Point(124, connectionTypeLabel.Location.Y),
                DataBindings = { new Binding("Text", bindingSource, "ConnectionType", true, DataSourceUpdateMode.OnPropertyChanged) }
            };
            connectionTypeComboBox.Items.AddRange(new object[] { "Flanged", "Threaded", "Welded", "Socket" });
            propertiesPanel.Controls.Add(connectionTypeComboBox);

            //Add "Connection Properties" label and combobox
           Label connectionPropertiesLabel = new Label
           {
               Text = "Connection Properties",
               AutoSize = false,
               Location = new Point(4, connectionTypeLabel.Bottom + 10),
               Size = new Size(100, 30),
               Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
           };
            propertiesPanel.Controls.Add(connectionPropertiesLabel);

            ComboBox connectionPropertiesComboBox = new ComboBox
            {
                Name = "ConnectionPropertiesComboBox",
                Size = new Size(100, 20),
                Location = new Point(124, connectionPropertiesLabel.Location.Y),
                DataBindings = { new Binding("Text", bindingSource, "ConnectionProperties", true, DataSourceUpdateMode.OnPropertyChanged) }
            };
            // Add some example items (adjust based on your requirements)
            connectionPropertiesComboBox.Items.AddRange(new object[] { "Standard", "High Pressure", "Low Pressure" });

            propertiesPanel.Controls.Add(connectionPropertiesComboBox);
            // Add "Material" label and textbox

            Label connectionMaterialLabel = new Label
            {
                Text = "Connection Material",
                AutoSize = false,
                Location = new Point(4, connectionPropertiesLabel.Bottom + 10),
                Size = new Size(100, 30),
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(connectionMaterialLabel);

            TextBox materialTextBox = new TextBox
            {
                Name = "MaterialTextBox",
                Size = new Size(100, 20),
                Location = new Point(124, connectionMaterialLabel.Location.Y)
            };

            var materialBinding = materialTextBox.DataBindings.Add(
                "Text",
                bindingSource,
                "ConnectionMaterial",
                true,
                DataSourceUpdateMode.OnPropertyChanged,
                string.Empty // nullValue -> empty string
            );
            propertiesPanel.Controls.Add(materialTextBox);

            // Add "Thickness" label and textbox
            Label thicknessLabel = new Label
            {
                Text = "Thickness",
                Location = new Point(4, connectionMaterialLabel.Bottom + 10),
                Size = new Size(100, 20),
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(thicknessLabel);

            TextBox thicknessTextBox = new TextBox
            {
                Name = "ThicknessTextBox",
                Size = new Size(100, thicknessLabel.Location.Y),
                Location = new Point(124, thicknessLabel.Location.Y),
            };

            var thicknessBinding = thicknessTextBox.DataBindings.Add(
                "Text",
                bindingSource,
                "NeckThicknessMeters",
                true,
                DataSourceUpdateMode.OnValidation);

            // model (meters) -> UI (mm)
            thicknessBinding.Format += (s, e) =>
            {
                if (e.Value is double meters)
                    e.Value = (meters * 1000.0).ToString(CultureInfo.CurrentCulture);
                else
                    e.Value = string.Empty;
            };

            // UI (mm) -> model (meters)
            thicknessBinding.Parse += (s, e) =>
            {
                string text = (e.Value ?? string.Empty).ToString().Trim();
                if (string.IsNullOrEmpty(text))
                {
                    // keep existing model value when input is empty
                    var current = bindingSource.Current as NozzleConfiguration;
                    e.Value = current?.NeckThicknessMeters ?? 0.0;
                    return;
                }

                if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double mm))
                {
                    e.Value = mm / 1000.0; // convert mm -> meters
                }
                else
                {
                    // invalid input: fallback to current model value
                    var current = bindingSource.Current as NozzleConfiguration;
                    e.Value = current?.NeckThicknessMeters ?? 0.0;
                }
            };

            propertiesPanel.Controls.Add(thicknessTextBox);

            // Add "Material" label and textbox
            Label neckMaterialLabel = new Label
            {
                Text = "Neck Material",
                Location = new Point(4, thicknessLabel.Bottom + 10),
                Size = new Size(100, 20),
                Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(neckMaterialLabel);

            TextBox neckMaterialTextBox = new TextBox
            {
                Name = "MaterialTextBox",
                Size = new Size(100, 20),
                Location = new Point(124, neckMaterialLabel.Location.Y),
            };
            propertiesPanel.Controls.Add(neckMaterialTextBox);

            var neckMaterialBinding = neckMaterialTextBox.DataBindings.Add(
                "Text",
                bindingSource,
                "NeckMaterial",
                true,
                DataSourceUpdateMode.OnPropertyChanged,
                string.Empty // nullValue -> empty string
            );

            // optional: trim whitespace before writing to model
            neckMaterialBinding.Parse += (s, e) =>
            {
                e.Value = (e.Value ?? "").ToString().Trim();
            };

            sketch.Controls.Add(propertiesPanel);
            propertiesPanel.BringToFront();

            sketch.SetVisualizationsEnabled(false);
        }

        public event EventHandler NewNozzleButtonClicked;
    }
}

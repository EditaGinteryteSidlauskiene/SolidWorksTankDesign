using SolidWorks.Interop.sldworks;
using SolidWorksTankDesign.MVP.Enums;
using SolidWorksTankDesign.TankSiteConfigurations;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SolidWorksTankDesign.MVP.Views.Controls
{
    /// <summary>
    /// Interactive control that visualizes a compartment cross-section and allows the user
    /// to select a nozzle position reference (left dished end, right dished end, or another nozzle)
    /// and specify the distance from that reference.
    /// </summary>
    internal class NozzlePositionControl : UserControl
    {
        private CompartmentConfiguration _compartmentConfig;
        private DishedEndAlignment _rightDishedEndAlignment;
        private NozzleConfiguration _currentNozzleConfiguration;

        /// <summary>The currently active reference type selected by the user via hotspot click.</summary>
        private NozzleReferenceType? _activeNozzleReferenceType = null;
        /// <summary>Tracks the previous click's reference type to detect consecutive OtherNozzle clicks for direction toggle.</summary>
        private NozzleReferenceType? _previousNozzleReferenceType = null;

        private PictureBox _positionPictureBox;

        // ===== Geometry constants =====
        private const int RectangleWidth = 150;
        private const int RectangleHeight = 75;
        private const int DishedEndHeight = 60;
        private const float CompartmentLengthLineWidth = 2f;
        private const int CompartmentLengthLineOffset = 20;
        private const float LineArrowLength = 6f;
        private const float LineArrowWidth = 4f;
        private const float DotRadius = 3.5f;
        private const int DistanceLineLength = 40;
        private const int DistanceTextBoxYOffset = 30;

        List<Hotspot> _hotspots = new List<Hotspot>();
        TextBox _distanceTextBox;
        ComboBox _referenceNozzleComboBox;
        Label _compartmentLengthLabel;
        /// <summary>Timer for the validation error flash effect. Disposed on tick or control disposal.</summary>
        Timer _flashTimer;
        /// <summary>When OtherNozzle is selected, indicates whether the reference arrow points left or right.</summary>
        bool _isReferenceToLeft = false;
        
        /// <summary>
        /// Initializes the control with the compartment and nozzle configuration data.
        /// </summary>
        /// <param name="compartmentConfig">The compartment this nozzle belongs to.</param>
        /// <param name="currentNozzleConfiguration">The nozzle being positioned.</param>
        /// <param name="rightDishedEndAlignment">Alignment of the right dished end (from the next compartment).</param>
        public NozzlePositionControl(
            CompartmentConfiguration compartmentConfig,
            NozzleConfiguration currentNozzleConfiguration,
            DishedEndAlignment rightDishedEndAlignment)
        {
            _compartmentConfig = compartmentConfig ?? throw new ArgumentNullException(nameof(compartmentConfig));
            _currentNozzleConfiguration = currentNozzleConfiguration ?? throw new ArgumentNullException(nameof(currentNozzleConfiguration));
            _rightDishedEndAlignment = rightDishedEndAlignment;
            InitializeComponent();
        }

        /// <summary>
        /// Creates and configures all child controls: PictureBox for drawing, distance TextBox
        /// with data binding (mm ↔ meters), and reference nozzle ComboBox with placeholder.
        /// </summary>
        private void InitializeComponent()
        {
            _positionPictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Transparent
            };
            Controls.Add(_positionPictureBox);

            _positionPictureBox.Paint += _positionPictureBox_Paint;
            _positionPictureBox.MouseClick += _positionPictureBox_MouseClick;

            // Distance input TextBox — hidden until a reference hotspot is clicked
            _distanceTextBox = new TextBox
            {
                Name = "DistanceTextBox",
                Width = DistanceLineLength,
                Height = 15,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                Visible = false
            };

            var distanceBinding = _distanceTextBox.DataBindings.Add(
                "Text",
                _currentNozzleConfiguration,
                "DistanceFromReference",
                true,
                DataSourceUpdateMode.OnPropertyChanged);

            // model (meters) → UI (mm)
            distanceBinding.Format += (s, ev) =>
            {
                if (ev.Value is double meters)
                    ev.Value = (meters * 1000.0).ToString(System.Globalization.CultureInfo.CurrentCulture);
                else
                    ev.Value = string.Empty;
            };

            // UI (mm) → model (meters): validates input before writing to model
            distanceBinding.Parse += (s, ev) =>
            {
                string text = (ev.Value ?? string.Empty).ToString().Trim();

                // Empty input — keep existing model value
                if (string.IsNullOrEmpty(text))
                {
                    ev.Value = _currentNozzleConfiguration.DistanceFromReference;
                    return;
                }

                if (double.TryParse(text,
                    System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                    System.Globalization.CultureInfo.CurrentCulture,
                    out double mm))
                {
                    // Reject negative distances
                    if (mm < 0)
                    {
                        FlashTextBoxError(_distanceTextBox);
                        ev.Value = _currentNozzleConfiguration.DistanceFromReference;
                        return;
                    }

                    // Reject distances exceeding compartment length
                    double maxMm = _compartmentConfig.Length;
                    if (mm > maxMm)
                    {
                        FlashTextBoxError(_distanceTextBox);
                        ev.Value = _currentNozzleConfiguration.DistanceFromReference;
                        return;
                    }

                    // Valid input — clear any error highlight and convert mm → meters
                    _distanceTextBox.BackColor = Color.White;
                    ev.Value = mm / 1000.0;
                }
                else
                {
                    // Non-numeric input — revert to current model value
                    ev.Value = _currentNozzleConfiguration.DistanceFromReference;
                }
            };

            _distanceTextBox.GotFocus += _distanceTextBox_GotFocus;
            Controls.Add(_distanceTextBox);
            _distanceTextBox.BringToFront();

            // Build the reference nozzle list: all nozzles in compartment except the current one
            List<NozzleConfiguration> nozzleConfigurations = _compartmentConfig.NozzleConfigurations.ToList();
            nozzleConfigurations.Remove(_currentNozzleConfiguration);

            // Insert a placeholder at index 0 so "Select" shows by default
            var placeholder = new NozzleConfiguration { Designation = "Select" };
            nozzleConfigurations.Insert(0, placeholder);

            _referenceNozzleComboBox = new ComboBox
            {
                Name = "ReferenceNozzleComboBox",
                Size = new Size(55, 20),
                DropDownStyle = ComboBoxStyle.DropDownList,
                DataSource = nozzleConfigurations,
                DisplayMember = "Designation",
                ValueMember = "Id",
                Visible = false
            };
            _referenceNozzleComboBox.SelectedIndexChanged += (s, ev) =>
            {
                // Real nozzle selected — save its Id to the model
                if (_referenceNozzleComboBox.SelectedIndex > 0
                    && _referenceNozzleComboBox.SelectedValue is Guid selectedId)
                {
                    _currentNozzleConfiguration.ReferenceNozzleId = selectedId;
                }
                // Placeholder "Select" chosen — clear any previously saved reference
                else if (_referenceNozzleComboBox.SelectedIndex == 0)
                {
                    _currentNozzleConfiguration.ReferenceNozzleId = null;
                }
            };

            Controls.Add(_referenceNozzleComboBox);
            _referenceNozzleComboBox.BringToFront();

        }

        /// <summary>
        /// Selects all text when the distance TextBox receives focus.
        /// Uses BeginInvoke so SelectAll runs after the default focus/caret handling.
        /// </summary>
        private void _distanceTextBox_GotFocus(object sender, EventArgs e)
        {
            _distanceTextBox.BeginInvoke(new Action(() => _distanceTextBox.SelectAll()));
        }

        /// <summary>
        /// Briefly sets the TextBox background to MistyRose as visual validation feedback.
        /// Resets to white after 800ms only if the current text is valid.
        /// Cancels any previously running flash timer to prevent overlapping flashes.
        /// </summary>
        private void FlashTextBoxError(TextBox tb)
        {
            // Cancel any previously running flash to prevent overlapping timers
            if (_flashTimer != null)
            {
                _flashTimer.Stop();
                _flashTimer.Dispose();
                _flashTimer = null;
            }

            // Immediately show error highlight
            tb.BackColor = Color.MistyRose;

            // After 800ms, reset to white only if the user has corrected the input
            _flashTimer = new Timer { Interval = 800 };
            _flashTimer.Tick += (s, ev) =>
            {
                if (IsDistanceValid(tb.Text))
                    tb.BackColor = Color.White;

                _flashTimer.Stop();
                _flashTimer.Dispose();
                _flashTimer = null;
            };
            _flashTimer.Start();
        }

        /// <summary>
        /// Validates that the given text represents a non-negative number
        /// not exceeding the compartment length (in mm).
        /// </summary>
        private bool IsDistanceValid(string text)
        {
            text = (text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(text))
                return true;

            if (!double.TryParse(text,
                System.Globalization.NumberStyles.Float | System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.CurrentCulture,
                out double mm))
                return false;

            return mm >= 0 && mm <= _compartmentConfig.Length;
        }

        /// <summary>
        /// Handles mouse clicks on the compartment visualization.
        /// Converts click position to normalized coordinates, hit-tests against hotspots,
        /// updates the active reference type, and toggles direction for consecutive OtherNozzle clicks.
        /// Clears ReferenceNozzleId when switching away from OtherNozzle.
        /// </summary>
        private void _positionPictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            // Convert pixel click position to normalized 0..1 coordinates
            float clickX = e.Location.X / (float)_positionPictureBox.Width;
            float clickY = e.Location.Y / (float)_positionPictureBox.Height;

            for (int i = 0; i < _hotspots.Count; i++)
            {
                // Compute tolerance in normalized space (elliptical hit area)
                float horizontalTolerance = (_hotspots[i].Tolerance) / _positionPictureBox.Width;
                float verticalTolerance = (_hotspots[i].Tolerance) / _positionPictureBox.Height;

                float dx = clickX - _hotspots[i].X;
                float dy = clickY - _hotspots[i].Y;
                float nx = dx / horizontalTolerance;
                float ny = dy / verticalTolerance;

                // Check if click is within the elliptical tolerance area
                if ((nx * nx + ny * ny) <= 1f)
                {
                    _activeNozzleReferenceType = _hotspots[i].NozzleReferenceType;

                    // Toggle direction only on consecutive OtherNozzle clicks
                    if (_activeNozzleReferenceType == NozzleReferenceType.OtherNozzle
                        && _previousNozzleReferenceType == NozzleReferenceType.OtherNozzle)
                    {
                        _isReferenceToLeft = !_isReferenceToLeft;
                    }
                    // Switching away from OtherNozzle — clear stale reference
                    else if (_activeNozzleReferenceType != NozzleReferenceType.OtherNozzle)
                    {
                        _currentNozzleConfiguration.ReferenceNozzleId = null;
                    }

                    // Persist to model and trigger repaint
                    _currentNozzleConfiguration.ReferenceType = _activeNozzleReferenceType.Value;
                    _previousNozzleReferenceType = _activeNozzleReferenceType;
                    _positionPictureBox.Invalidate();

                    break;
                }
            }
        }

        /// <summary>
        /// Draws the distance measurement line with arrowheads for the active reference type.
        /// For OtherNozzle, also positions the reference nozzle ComboBox.
        /// Returns the line start point for positioning the distance TextBox.
        /// </summary>
        private Point DrawHotspotVisualization(Graphics graphics, Rectangle rectangle)
        {
            int lineStartX = 0;
            int lineEndX = 0;
            bool leftArrowHead = false;
            bool rightArrowHead = false;

            // Determine line position and arrowhead direction based on active reference type
            if (_activeNozzleReferenceType == NozzleReferenceType.LeftDishedEnd)
            {
                // Arrow starts at left edge, points right
                lineStartX = rectangle.Left;
                lineEndX = lineStartX + DistanceLineLength;
                rightArrowHead = true;
            }
                
            else if (_activeNozzleReferenceType == NozzleReferenceType.RightDishedEnd)
            {
                // Arrow ends at right edge, points left
                lineStartX = rectangle.Right - DistanceLineLength;
                lineEndX = rectangle.Right;
                leftArrowHead = true;
            }

            else if (_activeNozzleReferenceType == NozzleReferenceType.OtherNozzle)
            {
                if(_isReferenceToLeft)
                {
                    // Arrow points left from nozzle center; ComboBox appears to the right
                    lineStartX = rectangle.Location.X + rectangle.Width / 2 - DistanceLineLength;
                    lineEndX = lineStartX + DistanceLineLength;
                    leftArrowHead = true;

                    PositionReferenceNozzleComboBox(new Point(lineEndX + 10, rectangle.Location.Y - 10 - DistanceTextBoxYOffset));
                }
                else
                {
                    // Arrow points right from nozzle center; ComboBox appears to the left
                    lineStartX = rectangle.Location.X + rectangle.Width / 2;
                    lineEndX = lineStartX + DistanceLineLength;
                    rightArrowHead = true;

                    PositionReferenceNozzleComboBox(new Point(lineStartX - 10 - _referenceNozzleComboBox.Width, rectangle.Location.Y - 10 - DistanceTextBoxYOffset));
                }
            }


            PaintLineWithArrowheads(
                graphics,
                lineStartX, lineEndX,
                rectangle.Location.Y - 10,
                leftArrowHead,
                rightArrowHead);

            leftArrowHead = false;
            rightArrowHead = false;

            return new Point(lineStartX, rectangle.Location.Y - 10);
        }

        /// <summary>
        /// Shows the reference nozzle ComboBox at the specified location.
        /// </summary>
        private void PositionReferenceNozzleComboBox(Point comboBoxLocation)
        {
            _referenceNozzleComboBox.Visible = true;
            _referenceNozzleComboBox.Location = comboBoxLocation;
        }

        /// <summary>
        /// Main paint handler. Draws the compartment rectangle, dished ends, length visualization,
        /// clickable hotspot dots, and (when active) the distance measurement line with TextBox.
        /// Hides the reference nozzle ComboBox at the start of each paint cycle;
        /// it is re-shown by DrawHotspotVisualization only when OtherNozzle is active.
        /// </summary>
        private void _positionPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Hide ComboBox at the start — DrawHotspotVisualization re-shows it only for OtherNozzle
            _referenceNozzleComboBox.Visible = false;

            Graphics graphics = e.Graphics;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            Rectangle rectangle = GetCompartmentRectangle();

            // Draw static compartment elements
            PaintRectangle(graphics, rectangle);
            PaintDishedEnds(graphics, rectangle);
            AddCompartmentLengthVisualization(graphics, rectangle);

            // Rebuild hotspots each paint cycle (positions depend on current layout)
            _hotspots.Clear();
            AddCornerDots(graphics, rectangle);

            // Only show OtherNozzle hotspot when there are multiple nozzles to reference
            if (_compartmentConfig.NozzleConfigurations.Count > 1)
                AddReferenceNozzleVisualization(graphics, rectangle);

            // Draw active distance visualization and position the TextBox
            if (_activeNozzleReferenceType.HasValue)
            {
                Point visualizationLocation = DrawHotspotVisualization(graphics, rectangle);
                PositionDistanceTextBox(visualizationLocation);
            }
        }

        /// <summary>
        /// Shows and positions the distance TextBox above the visualization line.
        /// </summary>
        private void PositionDistanceTextBox(Point visualizationLocation)
        {
            _distanceTextBox.Visible = true;
            _distanceTextBox.Location = new Point(visualizationLocation.X, visualizationLocation.Y - DistanceTextBoxYOffset);
        }

        /// <summary>
        /// Draws a small nozzle rectangle at the top center of the compartment with a clickable dot.
        /// Only shown when more than one nozzle exists in the compartment.
        /// </summary>
        private void AddReferenceNozzleVisualization(Graphics graphics, Rectangle rectangle)
        {
            int rectWidth = 13;
            int rectHeight = 20;

            // Center the nozzle rectangle on the top edge of the compartment
            Rectangle nozzleRectangle = new Rectangle(
                rectangle.X + rectangle.Width / 2 - rectWidth / 2, rectangle.Top - rectHeight / 2,
                rectWidth, rectHeight);

            PaintRectangle(graphics, nozzleRectangle);

            // Place the clickable dot at the center of the nozzle rectangle
            Point nozzlePointLocation = new Point(nozzleRectangle.X + rectWidth / 2,
                nozzleRectangle.Y + rectHeight / 2);
            DrawDot(graphics, nozzlePointLocation);
            AddHotspot(nozzlePointLocation, NozzleReferenceType.OtherNozzle);
        }

        /// <summary>
        /// Draws clickable dots on the upper-left and upper-right corners of the compartment rectangle,
        /// representing the left and right dished end reference points.
        /// </summary>
        private void AddCornerDots(Graphics graphics, Rectangle rectangle)
        {
            // Draw dots on upper corners
            Point dotLocation = new Point(rectangle.Left, rectangle.Top);
            DrawDot(graphics, dotLocation);
            AddHotspot(dotLocation, NozzleReferenceType.LeftDishedEnd);

            dotLocation = new Point(rectangle.Right, rectangle.Top);
            DrawDot(graphics, dotLocation);
            AddHotspot(dotLocation, NozzleReferenceType.RightDishedEnd);
        }

        /// <summary>
        /// Draws a filled and outlined circle at the given location.
        /// </summary>
        private void DrawDot(Graphics graphics, Point dotLocation)
        {
            RectangleF dotRectangle = new RectangleF(
                dotLocation.X - DotRadius, dotLocation.Y - DotRadius,
                DotRadius * 2, DotRadius * 2);

            using (Brush dotBursh = new SolidBrush(Color.DarkRed))
            using (Pen dotPen = new Pen(Color.DarkRed, 1f))
            {
                graphics.FillEllipse(dotBursh, dotRectangle);
                graphics.DrawEllipse(dotPen, dotRectangle);
            }
        }

        /// <summary>
        /// Registers a clickable hotspot at the given pixel location (converted to normalized 0..1 coordinates).
        /// </summary>
        private void AddHotspot(Point dotLocation, NozzleReferenceType nozzleReferenceType)
        {
            Hotspot hotspot = new Hotspot
            {
                X = dotLocation.X / (float)_positionPictureBox.Width,
                Y = dotLocation.Y / (float)_positionPictureBox.Height,
                NozzleReferenceType = nozzleReferenceType
            };

            _hotspots.Add(hotspot);
        }

        /// <summary>
        /// Draws the compartment length indicator: a horizontal line with arrowheads at both ends
        /// below the compartment rectangle, plus a label showing the length in mm.
        /// </summary>
        private void AddCompartmentLengthVisualization(Graphics graphics, Rectangle rectangle)
        {
            PaintLineWithArrowheads(
                graphics, 
                rectangle.Left, 
                rectangle.Right, 
                rectangle.Bottom + CompartmentLengthLineOffset,
                true, true);

            AddCompartmentLengthLabel(rectangle);
        }

        /// <summary>
        /// Creates the compartment length label once and repositions it on each paint.
        /// </summary>
        private void AddCompartmentLengthLabel(Rectangle rectangle)
        {
            int labelWidth = 80;

            // Create the label only once; subsequent paints just reposition it
            if (_compartmentLengthLabel == null)
            {
                _compartmentLengthLabel = new Label
                {
                    Text = $"{_compartmentConfig.Length.ToString()} mm",
                    TextAlign = ContentAlignment.TopCenter,
                    ForeColor = Color.ForestGreen,
                    Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0),
                    Width = labelWidth,
                    Height = 13
                };
                _compartmentLengthLabel.SendToBack();
                _positionPictureBox.Controls.Add(_compartmentLengthLabel);
            }

            _compartmentLengthLabel.Location = new Point(rectangle.X + RectangleWidth / 2 - labelWidth / 2, rectangle.Bottom + 2);
        }

        /// <summary>
        /// Draws a horizontal line with optional arrowheads at either end.
        /// When only one arrowhead is drawn, adds a perpendicular end cap on the opposite side.
        /// </summary>
        private void PaintLineWithArrowheads(
            Graphics graphics, 
            int startX, int endX, 
            int lineLocationY, 
            bool leftArrowHead, bool rightArrowHead)
        {
            using (Pen linePen = new Pen(Color.ForestGreen, CompartmentLengthLineWidth))
            {
                graphics.DrawLine(linePen, startX, lineLocationY, endX, lineLocationY);
            }

            // Draw arrowheads
            
            PointF leftWing1 = new PointF(startX + LineArrowLength,
                                          lineLocationY - LineArrowWidth);
            PointF leftWing2 = new PointF(startX + LineArrowLength,
                                      lineLocationY + LineArrowWidth);

            PointF rightWing1 = new PointF(endX - LineArrowLength,
                                      lineLocationY - LineArrowWidth);
            PointF rightWing2 = new PointF(endX - LineArrowLength,
                                      lineLocationY + LineArrowWidth);

            // Fill arrowhead triangles
            using (Brush brush = new SolidBrush(Color.ForestGreen))
            {
                if (leftArrowHead)
                    graphics.FillPolygon(brush, new PointF[] { new Point(startX, lineLocationY), leftWing1, leftWing2 });

                if (rightArrowHead)
                    graphics.FillPolygon(brush, new PointF[] { new Point(endX, lineLocationY), rightWing1, rightWing2 });
            }

            // When only one arrowhead is shown, draw a perpendicular end cap on the other side
            if (!(leftArrowHead && rightArrowHead))
            {
                using (Pen lineEnding = new Pen(Color.ForestGreen, CompartmentLengthLineWidth))
                {
                    if (leftArrowHead)
                        graphics.DrawLine(lineEnding, endX, lineLocationY - 5, endX, lineLocationY + 5);
                    else
                        graphics.DrawLine(lineEnding, startX, lineLocationY - 5, startX, lineLocationY + 5);
                }
            }
        }

        /// <summary>
        /// Computes the centered compartment rectangle within the PictureBox.
        /// </summary>
        private Rectangle GetCompartmentRectangle()
        {
            // Create rectangle
            int rectangleCenterX = _positionPictureBox.Width / 2 - RectangleWidth / 2;
            int rectangleCenterY = _positionPictureBox.Height / 2 - RectangleHeight / 3;

            return new Rectangle(rectangleCenterX, rectangleCenterY, RectangleWidth, RectangleHeight);
        }

        /// <summary>
        /// Draws dashed semi-circular arcs on both sides of the compartment rectangle
        /// representing the left and right dished ends.
        /// </summary>
        private void PaintDishedEnds(Graphics graphics, Rectangle rectangle)
        {
            int circleDiameter = rectangle.Height;
            using (Pen dishedEndPen = new Pen(Color.Black, 2f))
            {
                dishedEndPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                DrawDishedEnd(graphics, dishedEndPen, rectangle.Left, rectangle.Top, circleDiameter, _compartmentConfig.LeftDishedEndAlignment);
                DrawDishedEnd(graphics, dishedEndPen, rectangle.Right, rectangle.Top, circleDiameter, _rightDishedEndAlignment);
            }

        }

        /// <summary>
        /// Draws a single dished end arc at the specified edge, oriented left or right.
        /// </summary>
        private void DrawDishedEnd(Graphics graphics, Pen dishedEndPen, int edgeX, int topY, int circleDiameter, DishedEndAlignment dishedEndAlignment)
        {
            // Arc bounding rectangle centered on the edge
            Rectangle leftArcRect = new Rectangle(
                edgeX - DishedEndHeight / 2,
                topY,
                DishedEndHeight,
                circleDiameter);

            // Left-aligned: arc opens to the left (90° sweep); right-aligned: opens to the right (270° sweep)
            if (dishedEndAlignment == DishedEndAlignment.Left)
                graphics.DrawArc(dishedEndPen, leftArcRect, 90, 180);
            else
                graphics.DrawArc(dishedEndPen, leftArcRect, 270, 180);
        }

        /// <summary>
        /// Draws a black-outlined rectangle.
        /// </summary>
        private void PaintRectangle(Graphics graphics, Rectangle rectangle)
        {
            using (Pen rectanglePen = new Pen(Color.Black, 2f))
            {
                graphics.DrawRectangle(rectanglePen, rectangle);
            }
        }

        /// <summary>
        /// Disposes the flash timer to prevent it firing on a disposed control.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_flashTimer != null)
                {
                    _flashTimer.Stop();
                    _flashTimer.Dispose();
                    _flashTimer = null;
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Rebuilds the reference nozzle ComboBox data source from the current compartment's nozzle list,
        /// excluding the current nozzle. Restores the previous selection if ReferenceNozzleId is set.
        /// Called when the nozzle panel is expanded to pick up newly added nozzles.
        /// </summary>
        public void RefreshReferenceNozzleList()
        {
            // Build a fresh list excluding the current nozzle
            var nozzleConfigurations = _compartmentConfig.NozzleConfigurations
                .Where(n => n.Id != _currentNozzleConfiguration.Id)
                .ToList();

            // Insert placeholder so "Select" appears at index 0
            var placeholder = new NozzleConfiguration { Designation = "Select" };
            nozzleConfigurations.Insert(0, placeholder);

            // Replace the data source to pick up any newly added nozzles
            _referenceNozzleComboBox.DataSource = nozzleConfigurations;
            _referenceNozzleComboBox.DisplayMember = "Designation";
            _referenceNozzleComboBox.ValueMember = "Id";

            // Restore the previous selection if the referenced nozzle still exists
            if (_currentNozzleConfiguration.ReferenceNozzleId.HasValue)
            {
                for (int i = 0; i < nozzleConfigurations.Count; i++)
                {
                    if (nozzleConfigurations[i].Id == _currentNozzleConfiguration.ReferenceNozzleId.Value)
                    {
                        _referenceNozzleComboBox.SelectedIndex = i;
                        return;
                    }
                }
            }

            // No saved selection or referenced nozzle was deleted — show placeholder
            _referenceNozzleComboBox.SelectedIndex = 0;
        }
    }
}

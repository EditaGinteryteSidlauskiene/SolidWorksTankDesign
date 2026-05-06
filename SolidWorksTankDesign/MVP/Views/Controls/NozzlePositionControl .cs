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
        private const int CompartmentRectangleWidth = 152;
        private const int CompartmentRectangleHeight = 75;
        private const int NozzleRectangleWidth = 10;
        private const int NozzleRectangleHeight = 50;
        private const int DishedEndHeight = 40;
        private const float CompartmentLengthLineWidth = 2f;
        private const int CompartmentLengthLineOffset = 20;
        private const float LineArrowLength = 6f;
        private const float LineArrowWidth = 4f;
        private const float DotRadius = 3.5f;
        /// <summary>Length of the distance visualization line — always 1/4 of the compartment rectangle width.</summary>
        private const int DistanceLineLength = CompartmentRectangleWidth / 4;
        private const int DistanceTextBoxYOffset = 20;
        private const int ReferenceNozzleLabelYOffset = 60;

        /// <summary>Clickable hotspots rebuilt each paint cycle based on current layout.</summary>
        List<Hotspot> _hotspots = new List<Hotspot>();
        /// <summary>Distance input TextBox — hidden until a reference hotspot is clicked.</summary>
        TextBox _distanceTextBox;
        /// <summary>Label displaying the currently selected reference nozzle designation. Opens a context menu on click.</summary>
        Label _referenceNozzleLabel;
        /// <summary>Context menu listing all available reference nozzles (excludes the current nozzle).</summary>
        ContextMenuStrip _referenceNozzleMenu;
        /// <summary>Label showing the compartment length in mm below the compartment rectangle.</summary>
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
            InitializeComponents();
        }

        /// <summary>
        /// Creates and configures all child controls: PictureBox for drawing, distance TextBox
        /// with data binding (mm ↔ meters), and reference nozzle label with context menu.
        /// </summary>
        private void InitializeComponents()
        {
            // ===== PictureBox — main drawing surface =====
            _positionPictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.StretchImage,
                BackColor = Color.Transparent
            };
            Controls.Add(_positionPictureBox);

            _positionPictureBox.Paint += _positionPictureBox_Paint;
            _positionPictureBox.MouseClick += _positionPictureBox_MouseClick;
            _positionPictureBox.MouseMove += _positionPictureBox_MouseMove;

             // ===== Distance input TextBox — hidden until a reference hotspot is clicked =====
            _distanceTextBox = new TextBox
            {
                Name = "DistanceTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = DistanceLineLength,
                Height = 15,
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center,
                BackColor = this.BackColor,
                Cursor = Cursors.Hand,
                Visible = false
            };

            // Two-way data binding: model stores meters, UI displays mm
            var distanceBinding = _distanceTextBox.DataBindings.Add(
                "Text",
                _currentNozzleConfiguration,
                "DistanceFromReference",
                true,
                DataSourceUpdateMode.OnPropertyChanged);

            // Format: model (meters) → UI (mm)
            distanceBinding.Format += (s, ev) =>
            {
                if (ev.Value is double meters)
                    ev.Value = (meters * 1000.0).ToString(System.Globalization.CultureInfo.CurrentCulture);
                else
                    ev.Value = string.Empty;
            };

            // Parse: UI (mm) → model (meters) with validation
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
                    _distanceTextBox.BackColor = this.BackColor;
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

            // ===== Reference nozzle label + context menu =====
            // Build list of all nozzles in compartment except the current one
            List<NozzleConfiguration> nozzleConfigurations = _compartmentConfig.NozzleConfigurations.ToList();
            nozzleConfigurations.Remove(_currentNozzleConfiguration);

            // Build context menu with one item per available reference nozzle
            _referenceNozzleMenu = new ContextMenuStrip();
            foreach (var nozzle in nozzleConfigurations)
            {
                var item = _referenceNozzleMenu.Items.Add(nozzle.Designation);
                item.Tag = nozzle.Id;
                item.Click += (s, ev) =>
                {
                    var menuItem = s as ToolStripItem;
                    _referenceNozzleLabel.Text = menuItem.Text;
                    _currentNozzleConfiguration.ReferenceNozzleId = (Guid)menuItem.Tag;
                };
            }

            // Label that looks like plain text but opens the context menu on click
            _referenceNozzleLabel = new Label
            {
                Name = "ReferenceNozzleLabel",
                Size = new Size(55, 20),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font(Font.FontFamily, 9f, FontStyle.Bold),
                ForeColor = Color.Black,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
                Visible = false,
                Text = nozzleConfigurations.Count > 0 ? nozzleConfigurations[0].Designation : string.Empty
            };
            _referenceNozzleLabel.Click += (s, ev) =>
            {
                _referenceNozzleMenu.Show(_referenceNozzleLabel, new Point(0, _referenceNozzleLabel.Height));
            };

            Controls.Add(_referenceNozzleLabel);
            _referenceNozzleLabel.BringToFront();
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

            // After 800ms, reset to parent's background color only if the user has corrected the input
            _flashTimer = new Timer { Interval = 800 };
            _flashTimer.Tick += (s, ev) =>
            {
                if (IsDistanceValid(tb.Text))
                    tb.BackColor = this.BackColor;

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
        /// Changes the cursor to a hand when hovering over a clickable dot,
        /// and reverts to the default cursor when moving away.
        /// </summary>
        private void _positionPictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            float mouseX = e.Location.X / (float)_positionPictureBox.Width;
            float mouseY = e.Location.Y / (float)_positionPictureBox.Height;

            bool overDot = false;

            for (int i = 0; i < _hotspots.Count; i++)
            {
                float horizontalTolerance = _hotspots[i].Tolerance / _positionPictureBox.Width;
                float verticalTolerance = _hotspots[i].Tolerance / _positionPictureBox.Height;

                float dx = mouseX - _hotspots[i].X;
                float dy = mouseY - _hotspots[i].Y;
                float nx = dx / horizontalTolerance;
                float ny = dy / verticalTolerance;

                if ((nx * nx + ny * ny) <= 1f)
                {
                    overDot = true;
                    break;
                }
            }

            _positionPictureBox.Cursor = overDot ? Cursors.Hand : Cursors.Default;
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

                    // Persist reference type to model
                    _currentNozzleConfiguration.ReferenceType = _activeNozzleReferenceType.Value;
                    _previousNozzleReferenceType = _activeNozzleReferenceType;

                    // When OtherNozzle is activated, immediately persist the currently displayed
                    // label selection so the model is up to date even if the user doesn't
                    // open the menu (first nozzle already showing).
                    if (_activeNozzleReferenceType == NozzleReferenceType.OtherNozzle
                        && _referenceNozzleMenu.Items.Count > 0)
                    {
                        foreach (ToolStripItem item in _referenceNozzleMenu.Items)
                        {
                            if (item.Text == _referenceNozzleLabel.Text)
                            {
                                _currentNozzleConfiguration.ReferenceNozzleId = (Guid)item.Tag;
                                break;
                            }
                        }
                    }

                    _positionPictureBox.Invalidate();

                    break;
                }
            }
        }

        /// <summary>
        /// Draws the distance measurement line with arrowheads for the active reference type.
        /// For LeftDishedEnd and RightDishedEnd, the line is anchored to the corresponding edge.
        /// For OtherNozzle, the line extends from the nozzle center and the reference label is shown.
        /// Returns the line start point for positioning the distance TextBox.
        /// </summary>
        private Point DrawHotspotVisualization(Graphics graphics, Rectangle rectangle)
        {
            int lineStartX = 0;
            int lineEndX = 0;
            int nozzleVisualizationX = 0;

            // Determine line position and nozzle visualization placement based on active reference type
            if (_activeNozzleReferenceType == NozzleReferenceType.LeftDishedEnd)
            {
                // Arrow starts at left edge, points right toward the nozzle
                lineStartX = rectangle.Left;
                lineEndX = lineStartX + DistanceLineLength;
                nozzleVisualizationX = lineEndX;
            }
                
            else if (_activeNozzleReferenceType == NozzleReferenceType.RightDishedEnd)
            {
                // Arrow ends at right edge, points left toward the nozzle
                lineStartX = rectangle.Right - DistanceLineLength;
                lineEndX = rectangle.Right;
                nozzleVisualizationX = lineStartX;
            }

            else if (_activeNozzleReferenceType == NozzleReferenceType.OtherNozzle)
            {
                if(_isReferenceToLeft)
                {
                    // Arrow points left from nozzle center; reference label appears above the right end
                    lineStartX = rectangle.Location.X + rectangle.Width / 2 - DistanceLineLength;
                    lineEndX = lineStartX + DistanceLineLength;
                    nozzleVisualizationX = lineStartX;

                    PositionReferenceNozzleLabel(new Point(
                        lineEndX - _referenceNozzleLabel.Width / 2, 
                        rectangle.Location.Y - ReferenceNozzleLabelYOffset));
                }
                else
                {
                    // Arrow points right from nozzle center; reference label appears above the left end
                    lineStartX = rectangle.Location.X + rectangle.Width / 2;
                    lineEndX = lineStartX + DistanceLineLength;
                    nozzleVisualizationX = lineEndX;

                    PositionReferenceNozzleLabel(new Point(
                        lineStartX - _referenceNozzleLabel.Width / 2, 
                        rectangle.Location.Y - ReferenceNozzleLabelYOffset));
                }
            }

            // Draw the distance measurement line with arrowheads and dashed leader lines
            PaintLineWithArrowheads(
                graphics,
                lineStartX, lineEndX,
                rectangle.Location.Y - 20,
                rectangle.Location.Y - 30,
                30);

            // Draw a small nozzle rectangle at the end of the distance line
            PaintNozzleVisualization(graphics, nozzleVisualizationX, rectangle.Location.Y);

            return new Point(lineStartX, rectangle.Location.Y - 20);
        }

        /// <summary>
        /// Draws a small nozzle rectangle with a vertical center line at the specified position.
        /// Represents the nozzle being positioned within the compartment.
        /// </summary>
        /// <param name="graphics">The graphics surface to draw on.</param>
        /// <param name="rectX">X coordinate of the nozzle rectangle center.</param>
        /// <param name="rectCenterY">Top Y coordinate of the compartment rectangle, used to vertically position the nozzle.</param>
        private void PaintNozzleVisualization(Graphics graphics, int rectX, int rectCenterY)
        {
            Rectangle nozzleRect = new Rectangle(
                rectX - NozzleRectangleWidth / 2, 
                rectCenterY - NozzleRectangleHeight / 4,
                NozzleRectangleWidth,
                NozzleRectangleHeight);

            PaintRectangle(graphics, nozzleRect, 1.5f);
            
            // Vertical center line extending below the nozzle rectangle
            using (Pen centraLinePen = new Pen(Color.ForestGreen, CompartmentLengthLineWidth))
            {
                graphics.DrawLine(
                    centraLinePen, 
                    nozzleRect.X + NozzleRectangleWidth / 2, 
                    nozzleRect.Y,
                    nozzleRect.X + NozzleRectangleWidth / 2,
                    nozzleRect.Y + NozzleRectangleHeight + 10);
            }
        }

        /// <summary>
        /// Shows the reference nozzle label at the specified location.
        /// </summary>
        private void PositionReferenceNozzleLabel(Point labelLocation)
        {
            _referenceNozzleLabel.Visible = true;
            _referenceNozzleLabel.Location = labelLocation;
        }

        /// <summary>
        /// Main paint handler. Draws all compartment visualization elements in the following order:
        /// 1. Compartment rectangle outline
        /// 2. Left and right dished end arcs
        /// 3. Compartment length line with arrowheads and label
        /// 4. Clickable hotspot dots (left/right dished end corners, and center OtherNozzle dot)
        /// 5. Active distance visualization line with nozzle rectangle and TextBox (when a reference is selected)
        /// Hides the reference nozzle label at the start; re-shown only when OtherNozzle is active.
        /// </summary>
        private void _positionPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Hide label at the start — DrawHotspotVisualization re-shows it only for OtherNozzle
            _referenceNozzleLabel.Visible = false;

            Graphics graphics = e.Graphics;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            Rectangle rectangle = GetCompartmentRectangle();

            // ===== Draw static compartment elements =====
            PaintRectangle(graphics, rectangle, 2f);
            PaintDishedEnds(graphics, rectangle);
            AddCompartmentLengthVisualization(graphics, rectangle);

            // ===== Rebuild hotspots each paint cycle (positions depend on current layout) =====
            _hotspots.Clear();
            AddCornerDots(graphics, rectangle);

            // Only show OtherNozzle hotspot when there are multiple nozzles to reference
            if (_compartmentConfig.NozzleConfigurations.Count > 1)
                AddReferenceNozzleVisualization(graphics, rectangle);

            // ===== Draw active distance visualization and position the TextBox =====
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
        /// Draws a clickable dot at the top center of the compartment rectangle
        /// representing the OtherNozzle reference point.
        /// Only shown when more than one nozzle exists in the compartment.
        /// </summary>
        private void AddReferenceNozzleVisualization(Graphics graphics, Rectangle rectangle)
        {
            Point dotPointLocation = new Point(
                rectangle.X + rectangle.Width / 2,
                rectangle.Top);

            DrawDot(graphics, dotPointLocation);
            AddHotspot(dotPointLocation, NozzleReferenceType.OtherNozzle);
        }

        /// <summary>
        /// Draws clickable dots on the upper-left and upper-right corners of the compartment rectangle,
        /// representing the left and right dished end reference points.
        /// </summary>
        private void AddCornerDots(Graphics graphics, Rectangle rectangle)
        {
            // Left dished end dot (upper-left corner)
            Point dotLocation = new Point(rectangle.Left, rectangle.Top);
            DrawDot(graphics, dotLocation);
            AddHotspot(dotLocation, NozzleReferenceType.LeftDishedEnd);

            // Right dished end dot (upper-right corner)
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
        /// Registers a clickable hotspot at the given pixel location.
        /// Converts pixel coordinates to normalized 0..1 coordinates for resolution-independent hit-testing.
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
                rectangle.Bottom + CompartmentLengthLineOffset + 10,
                - CompartmentLengthLineOffset - 10);

            AddCompartmentLengthLabel(rectangle);
        }

        /// <summary>
        /// Creates the compartment length label once and repositions it on each paint.
        /// The label is added as a child of the PictureBox and sent to the back so it
        /// doesn't interfere with drawn elements.
        /// </summary>
        private void AddCompartmentLengthLabel(Rectangle rectangle)
        {
            int labelWidth = 80;

            // Create the label only once; subsequent paints just reposition it
            if (_compartmentLengthLabel == null)
            {
                _compartmentLengthLabel = new Label
                {
                    Text = _compartmentConfig.Length.ToString(),
                    TextAlign = ContentAlignment.TopCenter,
                    ForeColor = Color.ForestGreen,
                    Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0),
                    Width = labelWidth,
                    Height = 13
                };
                _compartmentLengthLabel.SendToBack();
                _positionPictureBox.Controls.Add(_compartmentLengthLabel);
            }

            // Center the label below the compartment rectangle
            _compartmentLengthLabel.Location = new Point(rectangle.X + CompartmentRectangleWidth / 2 - labelWidth / 2, rectangle.Bottom + 2);
        }

        /// <summary>
        /// Draws a horizontal line with arrowheads at both ends, plus vertical dashed leader lines
        /// extending from each endpoint to connect with the compartment rectangle edges.
        /// </summary>
        /// <param name="startX">Left endpoint X coordinate.</param>
        /// <param name="endX">Right endpoint X coordinate.</param>
        /// <param name="lineLocationY">Y coordinate of the horizontal line.</param>
        /// <param name="dashedLineY">Y coordinate where dashed leader lines originate.</param>
        /// <param name="dashedLineDistanceToObject">Length of dashed leader lines (positive = downward, negative = upward).</param>
        private void PaintLineWithArrowheads(
            Graphics graphics, 
            int startX, int endX, 
            int lineLocationY,
            int dashedLineY,
            int dashedLineDistanceToObject)
        {
            // Draw the main horizontal line
            using (Pen linePen = new Pen(Color.ForestGreen, CompartmentLengthLineWidth))
            {
                graphics.DrawLine(linePen, startX, lineLocationY, endX, lineLocationY);
            }

            // Compute arrowhead wing points
            PointF leftWing1 = new PointF(startX + LineArrowLength,
                                          lineLocationY - LineArrowWidth);
            PointF leftWing2 = new PointF(startX + LineArrowLength,
                                      lineLocationY + LineArrowWidth);

            PointF rightWing1 = new PointF(endX - LineArrowLength,
                                      lineLocationY - LineArrowWidth);
            PointF rightWing2 = new PointF(endX - LineArrowLength,
                                      lineLocationY + LineArrowWidth);

            // Fill arrowhead triangles at both ends
            using (Brush brush = new SolidBrush(Color.ForestGreen))
            {
                    graphics.FillPolygon(brush, new PointF[] { new Point(startX, lineLocationY), leftWing1, leftWing2 });
                    graphics.FillPolygon(brush, new PointF[] { new Point(endX, lineLocationY), rightWing1, rightWing2 });
            }

            // Draw vertical dashed leader lines connecting arrowheads to the compartment edges
            PaintDashedLines(graphics, startX, endX, dashedLineY, dashedLineDistanceToObject);
        }

        /// <summary>
        /// Draws vertical dashed leader lines at both endpoints of a measurement line.
        /// These connect the arrowhead line to the compartment rectangle edges.
        /// </summary>
        /// <param name="lineStartX">X coordinate of the left dashed line.</param>
        /// <param name="lineEndX">X coordinate of the right dashed line.</param>
        /// <param name="lineLocationY">Y coordinate where the dashed lines start.</param>
        /// <param name="distanceToObject">Length of each dashed line (positive = downward).</param>
        private void PaintDashedLines(Graphics graphics, int lineStartX, int lineEndX, int lineLocationY, int distanceToObject)
        {
            Point startPoint1 = new Point(lineStartX, lineLocationY);
            Point endPoint1 = new Point(lineStartX, lineLocationY + distanceToObject);

            Point startPoint2 = new Point(lineEndX, lineLocationY);
            Point endPoint2 = new Point(lineEndX, lineLocationY + distanceToObject);

            using (Pen dashedLinePen = new Pen(Color.ForestGreen, CompartmentLengthLineWidth))
            {
                dashedLinePen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                graphics.DrawLine(dashedLinePen, startPoint1, endPoint1);
                graphics.DrawLine(dashedLinePen, startPoint2, endPoint2);
            }
        }

        /// <summary>
        /// Computes the centered compartment rectangle within the PictureBox.
        /// The rectangle is offset slightly upward (by 1/3 of its height) to leave room
        /// for the length visualization below.
        /// </summary>
        private Rectangle GetCompartmentRectangle()
        {
            int rectangleCenterX = _positionPictureBox.Width / 2 - CompartmentRectangleWidth / 2;
            int rectangleCenterY = _positionPictureBox.Height / 2 - CompartmentRectangleHeight / 3;

            return new Rectangle(rectangleCenterX, rectangleCenterY, CompartmentRectangleWidth, CompartmentRectangleHeight);
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
        /// Draws a single dished end arc at the specified edge.
        /// Left-aligned arcs open to the left (convex outward); right-aligned arcs open to the right.
        /// </summary>
        /// <param name="edgeX">X coordinate of the compartment edge where the arc is centered.</param>
        /// <param name="topY">Top Y coordinate of the arc bounding rectangle.</param>
        /// <param name="circleDiameter">Diameter of the arc (matches compartment height).</param>
        /// <param name="dishedEndAlignment">Determines which direction the arc opens.</param>
        private void DrawDishedEnd(Graphics graphics, Pen dishedEndPen, int edgeX, int topY, int circleDiameter, DishedEndAlignment dishedEndAlignment)
        {
            // Arc bounding rectangle centered horizontally on the edge
            Rectangle leftArcRect = new Rectangle(
                edgeX - DishedEndHeight / 2,
                topY,
                DishedEndHeight,
                circleDiameter);

            // Left-aligned: arc sweeps from 90° for 180° (opens left)
            // Right-aligned: arc sweeps from 270° for 180° (opens right)
            if (dishedEndAlignment == DishedEndAlignment.Left)
                graphics.DrawArc(dishedEndPen, leftArcRect, 90, 180);
            else
                graphics.DrawArc(dishedEndPen, leftArcRect, 270, 180);
        }

        /// <summary>
        /// Draws a black-outlined rectangle with the specified line width.
        /// </summary>
        private void PaintRectangle(Graphics graphics, Rectangle rectangle, float lineWidth)
        {
            using (Pen rectanglePen = new Pen(Color.Black, lineWidth))
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
        /// Rebuilds the reference nozzle context menu from the current compartment's nozzle list,
        /// excluding the current nozzle. Restores the previous selection if ReferenceNozzleId is set.
        /// Called when the nozzle panel is expanded to pick up newly added nozzles.
        /// </summary>
        public void RefreshReferenceNozzleList()
        {
            // Build a fresh list excluding the current nozzle
            var nozzleConfigurations = _compartmentConfig.NozzleConfigurations
                .Where(n => n.Id != _currentNozzleConfiguration.Id)
                .ToList();

            // Rebuild context menu items
            _referenceNozzleMenu.Items.Clear();
            foreach (var nozzle in nozzleConfigurations)
            {
                var item = _referenceNozzleMenu.Items.Add(nozzle.Designation);
                item.Tag = nozzle.Id;
                item.Click += (s, ev) =>
                {
                    var menuItem = s as ToolStripItem;
                    _referenceNozzleLabel.Text = menuItem.Text;
                    _currentNozzleConfiguration.ReferenceNozzleId = (Guid)menuItem.Tag;
                };
            }

            // Restore the previous selection if the referenced nozzle still exists
            if (_currentNozzleConfiguration.ReferenceNozzleId.HasValue)
            {
                var match = nozzleConfigurations.FirstOrDefault(
                    n => n.Id == _currentNozzleConfiguration.ReferenceNozzleId.Value);
                if (match != null)
                {
                    _referenceNozzleLabel.Text = match.Designation;
                    return;
                }
            }

            // No saved selection or referenced nozzle was deleted — show first available
            if (nozzleConfigurations.Count > 0)
                _referenceNozzleLabel.Text = nozzleConfigurations[0].Designation;
        }
    }
}

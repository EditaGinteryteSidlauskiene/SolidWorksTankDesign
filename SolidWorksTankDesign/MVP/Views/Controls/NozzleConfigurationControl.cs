using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SolidWorksTankDesign.MVP.Enums;
using System.Globalization;
using System.Linq;
using SolidWorksTankDesign.TankSiteConfigurations;
using System.Diagnostics.Eventing.Reader;

namespace SolidWorksTankDesign.MVP.Views.Controls
{
    public class NozzleConfigurationControl : UserControl
    {
        private PictureBox _pictureBox;
        private List<Hotspot> _hotspots;
        private TextBox _offsetTextBox;
        private TextBox _nozzleCenterlineDistanceTextBox;
        private TextBox _nozzleBottomDistanceTextBox;
        private TextBox _rotationTextBox;
        private TextBox _lengthTextBox;
        private bool _showTankCenterlineDistance = false;
        private bool _showNozzleCenterlineDistance = false;
        private bool _showNozzleTopDistance = false;
        private bool _showNozzleMiddleDistance = false;
        private bool _showNozzleBottomDistance = false;
        private bool _isNozzleMiddleRectangleLong = false;
        private bool _showNozzleLength = false;
        private bool _isDiagonalCut = false;
        private string _lastActiveNozzleVisualization = null;
        private FlipDot _flipDot = FlipDot.Central;

        // Visualization tracking identifiers
        private const string VisualizationTop = "top";
        private const string VisualizationMiddle = "middle";
        private const string VisualizationBottom = "bottom";
        private bool _visualizationsEnabled = true;

        // ===== Geometry constants (normalized 0..1 coordinates unless noted) =====
        private const float NozzleCenterNormX = 0.45f;
        private const float NozzleLineBaseX = 0.3f;
        private const float NozzleLineStartY = 0.11f;
        private const float NozzleLineEndY = 0.96f;
        private const float NozzleRectWidth = 0.08f;
        private const float DefaultRectHeight = 0.3f;
        private const float InnerCircleRadius = 0.34f;
        private const float OuterCircleRadius = 0.37f;
        private const float DotRadiusPx = 3.5f;
        private const float Dot3InsetPx = 20f;
        private const float DotInsetPx = 28f;
        private const float DotInsetExtra = 0.01f;
        private const float RectOffsetPx = 20f;
        private const float LineExtensionPx = 20f;
        private const float CenterLineOverflowPx = 31;
        private const float CenterLineTopOverflowPx = 55;
        private const float DistVisualizationOffsetNorm = 0.3f;
        private const float DistanceLineY = 1.04f;
        private const float DistLinePenWidth = 1.7f;
        private const float VertDistPerpOffset = 0.1f;
        private const float DistPerpOffsetNorm = 0.07f;

        /// <summary>Raised when the active distance reference type changes (e.g. tank centerline, nozzle centerline, top/middle/bottom).</summary>
        public event EventHandler<DistanceChangedEventArgs> DistanceChanged;
        /// <summary>Raised when the nozzle flip state is toggled.</summary>
        public event EventHandler<EventArgs> FlipStateChanged;
        

        /// <summary>
        /// Event arguments indicating which distance reference type was activated.
        /// Exactly one of <see cref="TopReferenceType"/> or <see cref="BottomReferencePoint"/> is set.
        /// </summary>
        public class DistanceChangedEventArgs : EventArgs
        {
            public NozzleTopReferenceType? TopReferenceType { get; set; }
            public NozzleBottomReferencePoint? BottomReferencePoint { get; set; }
        }

        public NozzleConfigurationControl()
        {
            InitializeComponents(); 
            _hotspots = new List<Hotspot>();
        }

        /// <summary>
        /// Direct reference to the NozzleConfiguration model for manual textbox ↔ model sync.
        /// </summary>
        private NozzleConfiguration _nozzleConfig;

        /// <summary>
        /// Initializes the control with the given NozzleConfiguration model.
        /// Populates all distance textboxes from the model (meters → mm).
        /// </summary>
        public void SetNozzleConfiguration(NozzleConfiguration nozzleConfig)
        {
            if (nozzleConfig == null) return;
            _nozzleConfig = nozzleConfig;

            UpdateTextBoxFromModel(_offsetTextBox);
            UpdateTextBoxFromModel(_nozzleCenterlineDistanceTextBox);
            UpdateTextBoxFromModel(_nozzleBottomDistanceTextBox);
            UpdateTextBoxFromModel(_rotationTextBox);

            _pictureBox?.Invalidate();
        }
        /// <summary>
        /// Creates and configures all child controls: PictureBox, distance/rotation TextBoxes, and Flip button.
        /// Positions are set dynamically during the Paint event.
        /// </summary>
        private void InitializeComponents()
        {
            this.Padding = Padding.Empty;
            _pictureBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.StretchImage
            };
            _pictureBox.MouseClick += PictureBox_MouseClick;
            _pictureBox.MouseMove += PictureBox_MouseMove;
            _pictureBox.Paint += PictureBox_Paint;
            Controls.Add(_pictureBox);

            // Add distance input TextBox
            _offsetTextBox = new TextBox
            {
                Name = "DistanceTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
            };

            _offsetTextBox.LostFocus += _distanceTextBox_LostFocus;
            _offsetTextBox.MouseClick += _distanceTextBox_MouseClick;

            // Position will be set dynamically in Paint event
            Controls.Add(_offsetTextBox);
            _offsetTextBox.BringToFront();  // ensure it's on top of PictureBox

            _nozzleCenterlineDistanceTextBox = new TextBox
            {
                Name = "CenterlineDistanceTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
                Visible = false
            };

            _nozzleCenterlineDistanceTextBox.LostFocus += _tankCenterlineDistanceTextBox_LostFocus;
            _nozzleCenterlineDistanceTextBox.MouseClick += _nozzleCenterlineDistanceTextBox_MouseClick;

            // Position will be set dynamically in Paint event
            Controls.Add(_nozzleCenterlineDistanceTextBox);
            _nozzleCenterlineDistanceTextBox.BringToFront();  // ensure it's on top of PictureBox

            _nozzleBottomDistanceTextBox = new TextBox
            {
                Name = "NozzleBottomDistanceTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Center,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
                Visible = false
            };
            _nozzleBottomDistanceTextBox.LostFocus += _nozzleBottomDistanceTextBox_LostFocus;
            _nozzleBottomDistanceTextBox.MouseClick += _nozzleBottomDistanceTextBox_MouseClick;

            // Position will be set dynamically in Paint event
            Controls.Add(_nozzleBottomDistanceTextBox);
            _nozzleBottomDistanceTextBox.BringToFront();

            _rotationTextBox = new TextBox
            {
                Name = "RotationTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = 25,
                Height = 15,
                BorderStyle = BorderStyle.None,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
                Visible = true
            };
            _rotationTextBox.LostFocus += _rotationTextBox_LostFocus;
            _rotationTextBox.MouseClick += _rotationTextBox_MouseClick;
            _rotationTextBox.PreviewKeyDown += _rotationTextBox_PreviewKeyDown;
            _rotationTextBox.KeyDown += _rotationTextBox_KeyDown;

            // Position will be set dynamically in Paint event
            Controls.Add(_rotationTextBox);
            _rotationTextBox.BringToFront();

            _lengthTextBox = new TextBox
            {
                Name = "LengthTextBox",
                Font = new Font(Font.FontFamily, 10f),
                Width = 25,
                Height = 15,
                BorderStyle = BorderStyle.None,
                BackColor = BackColor,
                Cursor = Cursors.Hand,
                Visible = false
            };
            _lengthTextBox.LostFocus += _lengthTextBox_LostFocus;
            _lengthTextBox.MouseClick += _lengthTextBox_MouseClick;
            _lengthTextBox.PreviewKeyDown += _lengthTextBox_PreviewKeyDown;
            _lengthTextBox.KeyDown += _lengthTextBox_KeyDown;

            // Position will be set dynamically in Paint event
            Controls.Add(_lengthTextBox);
            _lengthTextBox.BringToFront();
        }

        /// <summary>
        /// Commits the nozzle length value to the model when the length TextBox loses focus.
        /// </summary>
        private void _lengthTextBox_LostFocus(object sender, EventArgs e)
        {
            CommitTextBoxToModel(_lengthTextBox);
        }

        /// <summary>
        /// Selects all text in the length TextBox when clicked for easy overwriting.
        /// </summary>
        private void _lengthTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            _lengthTextBox.BeginInvoke(new Action(() => _lengthTextBox.SelectAll()));
        }

        /// <summary>
        /// Marks the Enter key as an input key so it can be handled in KeyDown.
        /// </summary>
        private void _lengthTextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                e.IsInputKey = true;
        }

        /// <summary>
        /// On Enter key, commits the nozzle length value to the model, refreshes the visualization, and re-selects the text.
        /// </summary>
        private void _lengthTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                CommitTextBoxToModel(_lengthTextBox);
                ToggleRotationVisualization();
                _lengthTextBox.SelectAll();
            }
        }


        /// <summary>
        /// Commits the rotation angle value to the model when the rotation TextBox loses focus.
        /// </summary>
        private void _rotationTextBox_LostFocus(object sender, EventArgs e)
        {
            CommitTextBoxToModel(_rotationTextBox);
        }

        /// <summary>
        /// Selects all text in the rotation TextBox when clicked for easy overwriting.
        /// </summary>
        private void _rotationTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            _rotationTextBox.BeginInvoke(new Action(() => _rotationTextBox.SelectAll()));
        }

        /// <summary>
        /// Marks the Enter key as an input key so it can be handled in KeyDown.
        /// </summary>
        private void _rotationTextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                e.IsInputKey = true;
        }

        /// <summary>
        /// On Enter key, commits the rotation value to the model, refreshes the visualization, and re-selects the text.
        /// </summary>
        private void _rotationTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                CommitTextBoxToModel(_rotationTextBox);
                ToggleRotationVisualization();
                _rotationTextBox.SelectAll();
            }
        }

        /// <summary>
        /// Commits the offset distance value to the model when the offset TextBox loses focus.
        /// </summary>
        private void _distanceTextBox_LostFocus(object sender, EventArgs e)
        {
            CommitTextBoxToModel(_offsetTextBox);
        }

        /// <summary>
        /// Selects all text in the offset distance TextBox when clicked for easy overwriting.
        /// </summary>
        private void _distanceTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            _offsetTextBox.BeginInvoke(new Action(() => _offsetTextBox.SelectAll()));
        }

        /// <summary>
        /// Commits the nozzle bottom distance value to the model when the TextBox loses focus.
        /// </summary>
        private void _nozzleBottomDistanceTextBox_LostFocus(object sender, EventArgs e)
        {
            CommitTextBoxToModel(_nozzleBottomDistanceTextBox);
        }

        /// <summary>
        /// Selects all text in the nozzle bottom distance TextBox when clicked for easy overwriting.
        /// </summary>
        private void _nozzleBottomDistanceTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            _nozzleBottomDistanceTextBox.BeginInvoke(new Action(() => _nozzleBottomDistanceTextBox.SelectAll()));
        }

        /// <summary>
        /// Commits the centerline distance value to the model when the TextBox loses focus.
        /// </summary>
        private void _tankCenterlineDistanceTextBox_LostFocus(object sender, EventArgs e)
        {
            CommitTextBoxToModel(_nozzleCenterlineDistanceTextBox);
        }

        /// <summary>
        /// Selects all text in the centerline distance TextBox when clicked for easy overwriting.
        /// </summary>
        private void _nozzleCenterlineDistanceTextBox_MouseClick(object sender, MouseEventArgs e)
        {
            _nozzleCenterlineDistanceTextBox.BeginInvoke(new Action(() => _nozzleCenterlineDistanceTextBox.SelectAll()));
        }

        /// <summary>
        /// On Enter key, commits the centerline distance to the model, repaints, and re-selects the text.
        /// </summary>
        private void _nozzleCenterlineDistanceTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                CommitTextBoxToModel(_nozzleCenterlineDistanceTextBox);
                _pictureBox.Invalidate();
                _nozzleCenterlineDistanceTextBox.SelectAll();
            }
        }

        /// <summary>
        /// On Enter key, commits the nozzle bottom distance to the model, repaints, and re-selects the text.
        /// </summary>
        private void _nozzleBottomDistanceTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                CommitTextBoxToModel(_nozzleBottomDistanceTextBox);
                _pictureBox.Invalidate();
                _nozzleBottomDistanceTextBox.SelectAll();
            }
        }


        /// <summary>
        /// Toggles the nozzle flip state. Called by the flip arrow hotspot click or the legacy flip button.
        /// </summary>
        public void ToggleFlip(FlipDot flipDot)
        {
            if (_flipDot == flipDot)
                return;

            _flipDot = flipDot;
            _pictureBox.Invalidate();
            FlipStateChanged?.Invoke(this, EventArgs.Empty);
        }
        /// <summary>
        /// Sets the background image for the PictureBox and triggers a repaint of all overlays.
        /// </summary>
        public void SetImage(Image image)
        {
            _pictureBox.Image = image;
            _pictureBox.Invalidate();  // ← triggers OnPaint to redraw overlays
        }
        /// <summary>
        /// Registers a clickable hotspot. Its position will be updated dynamically during Paint.
        /// </summary>
        public void AddHotspot(Hotspot hotspot)
        {
            _hotspots.Add(hotspot);
        }
        /// <summary>
        /// Computes the current dot positions in normalized (0..1) coordinates,
        /// using the same formulas as PictureBox_Paint so they stay in sync when flipped.
        /// Returns a list of (X, Y, Hotspot) tuples representing each clickable dot.
        /// </summary>
        private List<Tuple<float, float, Hotspot>> GetCurrentDotPositions()
        {
            var dots = new List<Tuple<float, float, Hotspot>>();

            // Read positions directly from the hotspot objects (updated by Paint)
            foreach (var hs in _hotspots)
            {
                if (hs != null)
                {
                    dots.Add(Tuple.Create(hs.X, hs.Y, hs));
                }
            }

            return dots;
        }

        /// <summary>
        /// Finds the first hotspot matching the given <see cref="NozzlePropertiesType"/>, or null if not found.
        /// </summary>
        private Hotspot FindHotspot(NozzlePropertiesType type)
        {
            foreach (var hs in _hotspots)
                if (hs.NozzlePropertiesType == type) return hs;
            return null;
        }

        /// <summary>
        /// Finds the first hotspot matching the given <see cref="NozzleTopReferenceType"/>, or null if not found.
        /// </summary>
        private Hotspot FindHotspotByRef(NozzleTopReferenceType type)
        {
            foreach (var hs in _hotspots)
                if (hs.ReferenceType == type) return hs;
            return null;
        }

        /// <summary>
        /// Finds the first hotspot matching the given <see cref="FlipDot"/> value, or null if not found.
        /// </summary>
        private Hotspot FindHotspotByFlipDot(FlipDot flipDot)
        {
            foreach (var hs in _hotspots)
                if (hs.FlipDot == flipDot) return hs;
            return null;
        }

        /// <summary>
        /// Finds the first hotspot matching the given <see cref="NozzleBottomReferencePoint"/>, or null if not found.
        /// </summary>
        private Hotspot FindHotspotByBottom(NozzleBottomReferencePoint type)
        {
            foreach (var hs in _hotspots)
                if (hs.BottomReferencePoint == type) return hs;
            return null;
        }

        /// <summary>
        /// Finds the first hotspot marked as a nozzle length hotspot, or null if not found.
        /// </summary>
        private Hotspot FindHotspotByNozzleLength()
        {
            foreach (var hs in _hotspots)
                if (hs.IsNozzleLength) return hs;
            return null;
        }

        /// <summary>
        /// Changes the cursor to a hand when hovering over a clickable dot,
        /// and reverts to the default cursor when moving away.
        /// </summary>
        private void PictureBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_visualizationsEnabled)
            {
                _pictureBox.Cursor = Cursors.Default;
                return;
            }

            Rectangle imageRect = GetImageRect();
            if (imageRect.IsEmpty || imageRect.Width == 0 || imageRect.Height == 0)
            {
                _pictureBox.Cursor = Cursors.Default;
                return;
            }

            float relativeX = (e.Location.X - imageRect.X) / (float)imageRect.Width;
            float relativeY = (e.Location.Y - imageRect.Y) / (float)imageRect.Height;

            if (_hotspots == null || _hotspots.Count == 0)
            {
                _pictureBox.Cursor = Cursors.Default;
                return;
            }

            var currentDots = GetCurrentDotPositions();
            bool overDot = false;

            foreach (var dot in currentDots)
            {
                Hotspot hotspot = dot.Item3;
                if (hotspot == null) continue;

                float tolerancePx = hotspot.Tolerance;
                float tolX = tolerancePx / imageRect.Width;
                float tolY = tolerancePx / imageRect.Height;

                float dx = relativeX - dot.Item1;
                float dy = relativeY - dot.Item2;
                float nx = dx / tolX;
                float ny = dy / tolY;

                if ((nx * nx + ny * ny) <= 1f)
                {
                    overDot = true;
                    break;
                }
            }

            _pictureBox.Cursor = overDot ? Cursors.Hand : Cursors.Default;
        }

        /// <summary>
        /// Handles mouse clicks on the PictureBox by converting the click position to normalized
        /// coordinates and performing hit-testing against all registered hotspots.
        /// Ignored when visualizations are disabled.
        /// </summary>
        private void PictureBox_MouseClick(object sender, MouseEventArgs e)
        {
            if (!_visualizationsEnabled) return;

            // Get the actual rectangle where the drawing is rendered
            Rectangle imageRect = GetImageRect();
            if (imageRect.IsEmpty || imageRect.Width == 0 || imageRect.Height == 0) return;

            // Convert to normalized (0..1) coordinates relative to imageRect
            float relativeX = (e.Location.X - imageRect.X) / (float)imageRect.Width;
            float relativeY = (e.Location.Y - imageRect.Y) / (float)imageRect.Height;

            if (_hotspots == null || _hotspots.Count == 0) return;

            // Hit-testing against dynamically computed dot positions (in sync with Paint)
            var currentDots = GetCurrentDotPositions();

            foreach (var dot in currentDots)
            {
                Hotspot hotspot = dot.Item3;
                if (hotspot == null) continue;

                float tolerancePx = hotspot.Tolerance;
                float tolX = tolerancePx / imageRect.Width;
                float tolY = tolerancePx / imageRect.Height;

                float dx = relativeX - dot.Item1;
                float dy = relativeY - dot.Item2;
                float nx = dx / tolX;
                float ny = dy / tolY;

                if ((nx * nx + ny * ny) <= 1f)
                {
                    if (hotspot.OnClick != null)
                    {
                        hotspot.OnClick.Invoke();
                    }
                    else
                    {
                        HandleHotspotClick(hotspot);
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// Routes a hotspot click to the appropriate visualization toggle method
        /// based on the hotspot's reference type or bottom reference point.
        /// </summary>
        private void HandleHotspotClick(Hotspot hotspot)
        {
            if (hotspot.ReferenceType == NozzleTopReferenceType.TankCenterline)
            {
                ToggleTankCenterlineDistanceVisualization();
            }
            else if (hotspot.ReferenceType == NozzleTopReferenceType.NozzleCenterline)
            {
                ToggleNozzleCenterlineDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Top)
            {
                ToggleNozzleTopDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Bottom)
            {
                ToggleNozzleBottomDistanceVisualization();
            }
            else if (hotspot.BottomReferencePoint == NozzleBottomReferencePoint.Middle)
            {
                ToggleNozzleMiddleDistanceVisualization();
            }
            else if (hotspot.FlipDot == FlipDot.Left)
            {
                ToggleFlip(FlipDot.Left);
            }
            else if (hotspot.FlipDot == FlipDot.Right)
            {
                ToggleFlip(FlipDot.Right);
            }
            else if (hotspot.FlipDot == FlipDot.Central)
            {
                ToggleFlip(FlipDot.Central);
            }
            else if (hotspot.IsNozzleLength)
            {
                ToggleNozzleLengthVisualization();
            }
        }

        /// <summary>
        /// The fixed drawing size matches the old effective area (control 400×400 minus 75px padding on each side).
        /// Keeping this constant ensures visualizations stay the same size regardless of the actual control dimensions,
        /// while the larger control provides overflow room for rotated content.
        /// </summary>
        private const int DrawingAreaSize = 250;

        /// <summary>
        /// Computes the rectangle (in PictureBox client coordinates) where the drawing area is located.
        /// When no image is loaded, returns a centered fixed-size rectangle matching the original
        /// effective drawing area. When an image is loaded, computes the letterboxed/pillarboxed rect.
        /// </summary>
        private const int DrawingAreaVerticalOffsetPx = 10;
        private const int DrawingAreaHorizontalOffsetPx = 6;

        private Rectangle GetImageRect()
        {
            if (_pictureBox.Image == null)
            {
                // Center a fixed-size drawing area within the PictureBox so that
                // visualizations render at the same size as before the padding was removed.
                int x = (_pictureBox.ClientSize.Width - DrawingAreaSize) / 2 + DrawingAreaHorizontalOffsetPx;
                int y = (_pictureBox.ClientSize.Height - DrawingAreaSize) / 2 - DrawingAreaVerticalOffsetPx;
                return new Rectangle(x, y, DrawingAreaSize, DrawingAreaSize);
            }

            float imageAspect = (float)_pictureBox.Image.Width / _pictureBox.Image.Height;
            float boxAspect = (float)_pictureBox.ClientSize.Width / _pictureBox.ClientSize.Height;

            if (boxAspect > imageAspect)
            {
                int height = _pictureBox.ClientSize.Height;
                int width = (int)(height * imageAspect);
                int x = (_pictureBox.ClientSize.Width - width) / 2;
                return new Rectangle(x, 0, width, height);
            }
            else
            {
                int width = _pictureBox.ClientSize.Width;
                int height = (int)(width / imageAspect);
                int y = (_pictureBox.ClientSize.Height - height) / 2;
                return new Rectangle(0, y, width, height);
            }
        }

        /// <summary>
        /// Converts normalized (0..1) coordinates to PictureBox control-space pixel coordinates
        /// using the given image display rectangle.
        /// </summary>
        private PointF ImageToControlPoint(float relX, float relY, Rectangle imageRect)
        {
            float px = imageRect.X + relX * imageRect.Width;
            float py = imageRect.Y + relY * imageRect.Height;
            return new PointF(px, py);
        }

        /// <summary>
        /// Main Paint handler. Draws all nozzle visualization elements in the following order:
        /// 1. Inner and outer circles (rotated)
        /// 2. Dashed center lines (rotated)
        /// 3. Nozzle line and rectangle (rotated)
        /// 4. Clickable dots for connection, centerline, top/middle/bottom reference points (rotated)
        /// 5. Distance visualization lines with end caps (rotated)
        /// 6. Horizontal offset distance line (rotated)
        /// 7. Semi-circular rotation arrow (non-rotated)
        /// 8. Positions TextBoxes and Flip button (non-rotated, accounting for rotation offset)
        /// 9. Updates all hotspot positions for hit-testing
        /// Skipped entirely when visualizations are disabled.
        /// </summary>
        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            if (!_visualizationsEnabled) return;

            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Get the image rectangle (where the actual image is drawn)
            Rectangle imageRect = GetImageRect();

            if (imageRect.IsEmpty) return;  // no image loaded yet

            // ===== SAVE TRANSFORM =====
            System.Drawing.Drawing2D.Matrix originalTransform = g.Transform.Clone();

            // Define a simple line in normalized coordinates (0..1)
            float centerX = NozzleCenterNormX;
            float startX = _flipDot == FlipDot.Right ? (centerX + (centerX - NozzleLineBaseX)) : 
                _flipDot == FlipDot.Left ? NozzleLineBaseX :
                centerX;
            float startY = NozzleLineStartY;
            float endX = startX;
            float endY = NozzleLineEndY;

            // Convert to control pixel coordinates
            PointF start = ImageToControlPoint(startX, startY, imageRect);
            PointF end = ImageToControlPoint(endX, endY, imageRect);

            // Add a rectangle at the midpoint of the line
            // Determine rectangle size based on which reference point is active
            float rectCenterX = _flipDot == FlipDot.Right ? (centerX + (centerX - NozzleLineBaseX)) :
                _flipDot == FlipDot.Left ? NozzleLineBaseX :
                centerX;
            float rectCenterY;
            float rectWidth = NozzleRectWidth;
            // Default height; overridden below for long-rectangle cases (bottom ref or middle-long)
            float rectHeight = DefaultRectHeight;

            // Rectangle will be created after startY/endY are finalized

            // Circle centers (normalized coords)
            float circleCenterNormX = centerX;
            float circleCenterNormY = (startY + endY) / 2f;

            // Compute the base angle offset of the nozzle line from circle center
            float dx = rectCenterX - circleCenterNormX;

            // Compute line endpoints so they extend beyond the outer circle border.
            // Since the whole picture rotates via GDI+ transform, endpoints stay fixed.
            float dxPx = dx * imageRect.Width;
            float outerRadiusPx = OuterCircleRadius * imageRect.Width;
            float pivotOffsetXPx = dxPx;

            double effectiveR = (double)(outerRadiusPx + LineExtensionPx);
            double endpointDisc = effectiveR * effectiveR - (double)pivotOffsetXPx * pivotOffsetXPx;
            if (endpointDisc > 0)
            {
                double t = Math.Sqrt(endpointDisc);
                startY = circleCenterNormY - (float)(t / imageRect.Height);
            }

            // Set bottom endpoint to align with the horizontal offset distance line
            endY = DistanceLineY;

            // Recompute control pixel coordinates after adjusting startY/endY
            start = ImageToControlPoint(startX, startY, imageRect);
            end = ImageToControlPoint(endX, endY, imageRect);

            float baseStartY = startY;
            float baseEndY = endY;

            // Position rectangle center below the line start
            float rectOffsetNorm = RectOffsetPx / imageRect.Height;

            // Compute the fixed top edge (same for all rectangle sizes)
            float rectTopY = startY + rectOffsetNorm - rectHeight / 2;

            // For long rectangles (bottom ref or middle-long), override rectHeight
            // so that top stays fixed and bottom is midway between middle and bottom dots
            bool isLongRect = _showNozzleBottomDistance
                || (_showNozzleMiddleDistance && _isNozzleMiddleRectangleLong);
            if (isLongRect)
            {
                float dotInsetNormEarly = DotInsetPx / imageRect.Height;
                float dot6NormY = endY - dotInsetNormEarly;
                float rectBottomY = (circleCenterNormY + dot6NormY) / 2f;
                rectHeight = rectBottomY - rectTopY;
            }

            // Compute center from fixed top edge and final height
            rectCenterY = rectTopY + rectHeight / 2;
            PointF rectCenter = ImageToControlPoint(rectCenterX, rectCenterY, imageRect);
            float rectWidthPx = rectWidth * imageRect.Width;
            float rectHeightPx = rectHeight * imageRect.Height;
            RectangleF rect = new RectangleF(
                rectCenter.X - rectWidthPx / 2,
                rectCenter.Y - rectHeightPx / 2,
                rectWidthPx,
                rectHeightPx
            );

            // ===== Compute circle center pixel position =====
            PointF circleCenter = ImageToControlPoint(circleCenterNormX, circleCenterNormY, imageRect);
            float innerCircleRadiusPx = InnerCircleRadius * imageRect.Width;
            float outerCircleRadiusPx = OuterCircleRadius * imageRect.Width;

            // ===== APPLY ROTATION TRANSFORM (pivot around circle center) =====
            if (_nozzleConfig != null && _nozzleConfig.RotationAngleDegrees != 0)
            {
                g.TranslateTransform(circleCenter.X, circleCenter.Y);
                g.RotateTransform((float)_nozzleConfig.RotationAngleDegrees);
                g.TranslateTransform(-circleCenter.X, -circleCenter.Y);
            }

            // ===== ROTATED: Draw the nozzle circles =====
            RectangleF innerCircleRect = new RectangleF(
                circleCenter.X - innerCircleRadiusPx,
                circleCenter.Y - innerCircleRadiusPx,
                innerCircleRadiusPx * 2,
                innerCircleRadiusPx * 2);
            RectangleF outerCircleRect = new RectangleF(
                circleCenter.X - outerCircleRadiusPx,
                circleCenter.Y - outerCircleRadiusPx,
                outerCircleRadiusPx * 2,
                outerCircleRadiusPx * 2);
            using (Pen circlePen = new Pen(Color.Gray, 3f))
            {
                g.DrawEllipse(circlePen, innerCircleRect);
                g.DrawEllipse(circlePen, outerCircleRect);
            }

            // ===== ROTATED: Draw dashed center lines through circle center =====
            using (Pen dashedPen = new Pen(Color.Black, 1.5f))
            {
                dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                // Vertical center line extending beyond the outer circle
                PointF vTop = new PointF(circleCenter.X, circleCenter.Y - outerCircleRadiusPx - CenterLineTopOverflowPx);
                PointF vBottom = new PointF(circleCenter.X, circleCenter.Y + outerCircleRadiusPx + CenterLineOverflowPx);
                g.DrawLine(dashedPen, vTop, vBottom);

                // Horizontal center line extending beyond the outer circle
                PointF hLeft = new PointF(circleCenter.X - outerCircleRadiusPx - CenterLineOverflowPx, circleCenter.Y);
                PointF hRight = new PointF(circleCenter.X + outerCircleRadiusPx + CenterLineOverflowPx, circleCenter.Y);
                g.DrawLine(dashedPen, hLeft, hRight);
            }

            // ===== ROTATED DRAWING: line, rectangle, dots =====

            // Draw the main vertical line
            using (Pen pen = new Pen(Color.DarkGreen, 1.5f))
            {
                g.DrawLine(pen, start, end);
            }

            // Draw the rectangle (or diagonally cut rectangle)
            using (Pen rectPen = new Pen(Color.Gray, 2f))
            {
                if (_isDiagonalCut && _showNozzleTopDistance)
                {
                    // Draw a pentagon: one top corner is replaced by a diagonal cut.
                    // The cut removes a triangle from the top corner on the side closest to the circle center.
                    float cutSize = rectWidthPx * 0.6f;
                    float left = rect.Left;
                    float right = rect.Right;
                    float top = rect.Top;
                    float bottom = rect.Bottom;

                    PointF[] points;
                    // Cut the top-left corner (closer to center when flipped right)
                    points = new PointF[]
                    {
                        new PointF(left, top),   // left edge, below cut
                        new PointF(left, top),   // top edge, right of cut
                        new PointF(right, top),            // top-right
                        new PointF(right, bottom - cutSize),         // bottom-right
                        new PointF(left, bottom),          // bottom-left
                    };
                    g.DrawPolygon(rectPen, points);
                }
                else
                {
                    g.DrawRectangle(rectPen, Rectangle.Round(rect));
                }
            }

            // ===== DOTS (drawn on top of circles, still rotated) =====

            // --- Dot 1: Connection (middle of the top border of the rectangle) ---
            float dot1Y = rectCenterY - rectHeight / 2;
            DrawDot(g, ImageToControlPoint(rectCenterX, dot1Y, imageRect), Color.Black);

            // --- Dot 3: NozzleCenterline - where outer circle intersects nozzle center line (top) ---
            double outerDiscTop = (double)(outerCircleRadiusPx * outerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;
            float dot3Y = outerDiscTop > 0
                ? circleCenterNormY - (float)(Math.Sqrt(outerDiscTop) / imageRect.Height)
                : startY + Dot3InsetPx / imageRect.Height;
            float dot3X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot3X, dot3Y, imageRect), Color.FromArgb(80, 135, 138));

            // --- Dot 4: BottomReferencePoint.Top - where inner circle intersects nozzle center line (top) ---
            float dotInsetNorm = DotInsetPx / imageRect.Height;
            double innerDiscTop = (double)(innerCircleRadiusPx * innerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;
            float dot4OrigY = innerDiscTop > 0
                ? circleCenterNormY - (float)(Math.Sqrt(innerDiscTop) / imageRect.Height)
                : startY + dotInsetNorm + DotInsetExtra;
            float dot4Y = dot4OrigY;
            float dot4X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot4X, dot4Y, imageRect), Color.Purple);

            // --- Dot 5: BottomReferencePoint.Middle ---
            float dot5OrigY = circleCenterNormY;
            float dot5X = rectCenterX;
            float dot5Y = dot5OrigY;
            DrawDot(g, ImageToControlPoint(dot5X, dot5Y, imageRect), Color.Purple);

            // --- Dot 6: BottomReferencePoint.Bottom - where inner circle intersects nozzle center line ---
            double innerDiscBottom = (double)(innerCircleRadiusPx * innerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;
            float dot6Y = innerDiscBottom > 0
                ? circleCenterNormY + (float)(Math.Sqrt(innerDiscBottom) / imageRect.Height)
                : endY - dotInsetNorm - DotInsetExtra;
            float dot6OrigY = dot6Y;
            float dot6X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot6X, dot6Y, imageRect), Color.Purple);

            // ===== ROTATED: Distance visualization lines (drawn in rotated space) =====

            // --- Dot 7: TankCenterline - where vertical center line meets outer circle border (top) ---
            float dot7X = circleCenterNormX;
            float dot7Y = circleCenterNormY - outerCircleRadiusPx / imageRect.Height;
            DrawDot(g, ImageToControlPoint(dot7X, dot7Y, imageRect), Color.FromArgb(80, 135, 138));

            // --- Draw flip dots
            // Middle dot
            float flipMiddleDotX = circleCenterNormX;
            float flipMiddleDotY = circleCenterNormY + outerCircleRadiusPx / imageRect.Height;
            DrawDot(g, ImageToControlPoint(flipMiddleDotX, flipMiddleDotY, imageRect), Color.MediumBlue);

            // Left dot
            float flipLeftDotX = NozzleLineBaseX;
            float flipDotOffsetPX = (circleCenterNormX - flipLeftDotX) * imageRect.Width;
            float flipLeftDotY = circleCenterNormY +
                (float)Math.Sqrt(
                outerCircleRadiusPx * outerCircleRadiusPx -
                flipDotOffsetPX * flipDotOffsetPX) / 
                imageRect.Height;
            DrawDot(g, ImageToControlPoint(flipLeftDotX, flipLeftDotY, imageRect), Color.MediumBlue);

            // Right dot
            float flipRightDotX = circleCenterNormX + (circleCenterNormX - NozzleLineBaseX);
            float flipRightDotY = circleCenterNormY +
                (float)Math.Sqrt(
                outerCircleRadiusPx * outerCircleRadiusPx -
                flipDotOffsetPX * flipDotOffsetPX) /
                imageRect.Height;
            DrawDot(g, ImageToControlPoint(flipRightDotX, flipRightDotY, imageRect), Color.MediumBlue);

            // Nozzle length dot
            float nozzleLengthDotX = rectCenterX;
            float nozzleLengthDotY = rectCenterY + rectHeight / 2;
            DrawDot(g, ImageToControlPoint(nozzleLengthDotX, nozzleLengthDotY, imageRect), Color.Purple);

            // Compute dotX for tank centerline visualization — mirror the nozzle-side
            // visualization X across the circle center to place it on the opposite side.
            // When Central, use the same position as Left.
            float tcVisualizationOffset = 0.15f;
            float tcBaseX = _flipDot == FlipDot.Central ? NozzleLineBaseX : rectCenterX;
            float tcNozzleSideX = _flipDot == FlipDot.Right
                ? (tcBaseX + tcVisualizationOffset)
                : (tcBaseX - tcVisualizationOffset);
            float dotX = tcNozzleSideX + 2f * (circleCenterNormX - tcNozzleSideX);
            float dotY = circleCenterNormY;

            // TankCenterline hotspot position will be updated at the end of Paint
            // together with all other hotspots, using the rotated dot7 position.

            // Store normalized coords for textbox positioning (will be rotated later)
            float _tcDistStartX = 0, _tcDistStartY = 0, _tcDistEndY = 0;
            float _ncDistStartX = 0, _ncDistStartY = 0, _ncDistEndY = 0;
            float _nbDistStartX = 0, _nbDistStartY = 0, _nbDistEndY = 0;

            // Compute base (0° rotation) geometry for stable visualization positioning
            float baseRectTopY = baseStartY + RectOffsetPx / imageRect.Height - DefaultRectHeight / 2;
            float baseRectHeight = DefaultRectHeight;

            bool baseIsLongRect = _showNozzleBottomDistance
                || (_showNozzleMiddleDistance && _isNozzleMiddleRectangleLong);
            if (baseIsLongRect)
            {
                float baseDotInsetNorm = DotInsetPx / imageRect.Height;
                float baseDot6NormY = baseEndY - baseDotInsetNorm;
                float baseRectBottomY = (circleCenterNormY + baseDot6NormY) / 2f;
                baseRectHeight = baseRectBottomY - baseRectTopY;
            }
            float baseRectCenterY = baseRectTopY + baseRectHeight / 2;

            float baseDotInsetNormVal = DotInsetPx / imageRect.Height;

            double baseOuterDiscTop = (double)(outerCircleRadiusPx * outerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;

            double baseInnerDiscTop = (double)(innerCircleRadiusPx * innerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;
            float baseDot4OrigY = baseInnerDiscTop > 0
                ? circleCenterNormY - (float)(Math.Sqrt(baseInnerDiscTop) / imageRect.Height)
                : baseStartY + baseDotInsetNormVal + DotInsetExtra;
            float baseDot5OrigY = circleCenterNormY;
            float baseDot6OrigY;
            double baseInnerDiscBottom = (double)(innerCircleRadiusPx * innerCircleRadiusPx) - (double)pivotOffsetXPx * pivotOffsetXPx;
            baseDot6OrigY = baseInnerDiscBottom > 0
                ? circleCenterNormY + (float)(Math.Sqrt(baseInnerDiscBottom) / imageRect.Height)
                : baseEndY - baseDotInsetNormVal;

            float baseDot3Y;
            baseDot3Y = baseOuterDiscTop > 0
                ? circleCenterNormY - (float)(Math.Sqrt(baseOuterDiscTop) / imageRect.Height)
                : baseStartY + Dot3InsetPx / imageRect.Height;

            float distanceVisualizationOffset = DistVisualizationOffsetNorm;
            if (_flipDot == FlipDot.Central) distanceVisualizationOffset += 0.15f;

            // Store pre-rotation coords for nozzle centerline visualization using base values
            if (_showNozzleCenterlineDistance)
            {
                _ncDistStartX = dotX;
                _ncDistStartY = baseDot3Y;
                _ncDistEndY = baseRectCenterY - baseRectHeight / 2;
            }

            // Store pre-rotation coords for nozzle top/bottom/middle visualization using base values
            if (_showNozzleTopDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, distanceVisualizationOffset);
                _nbDistStartY = baseDot4OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }
            else if (_showNozzleBottomDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, distanceVisualizationOffset);
                _nbDistStartY = baseDot6OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }
            else if (_showNozzleMiddleDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, distanceVisualizationOffset);
                _nbDistStartY = baseDot5OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }
            else if (_showNozzleLength)
            {
                _nbDistStartX = FlipOffset(rectCenterX, distanceVisualizationOffset);
                _nbDistStartY = nozzleLengthDotY;
                _nbDistEndY = baseRectCenterY - baseRectHeight / 2;
            }

            // ===== ROTATED: Draw distance visualizations =====
            // Draw tank centerline distance visualization (rotated)
            // Uses fixed Y endpoints like other visualizations (no circle intersection)
            if (_showTankCenterlineDistance)
            {
                float tcLineX = dotX;
                float tcStartY = baseRectCenterY - baseRectHeight / 2;
                float tcEndY = circleCenterNormY - outerCircleRadiusPx / imageRect.Height;

                _tcDistStartX = tcLineX;
                _tcDistStartY = tcStartY;
                _tcDistEndY = tcEndY;

                DrawDistanceLine(g, imageRect, tcLineX, tcStartY, tcEndY,
                    rectCenterX, circleCenterNormX);
            }

            // Draw nozzle centerline distance visualization (rotated)
            if (_showNozzleCenterlineDistance)
            {
                DrawDistanceLine(g, imageRect, _ncDistStartX, _ncDistStartY, _ncDistEndY, rectCenterX, rectCenterX);
            }

            // Draw nozzle top/bottom/middle distance visualization (rotated)
            if (_showNozzleTopDistance || _showNozzleBottomDistance || _showNozzleMiddleDistance)
            {
                DrawDistanceLine(g, imageRect, _nbDistStartX, _nbDistStartY, _nbDistEndY, rectCenterX, rectCenterX);
            }

            if (_showNozzleLength)
            {
                DrawDistanceLine(g, imageRect, _nbDistStartX, _nbDistStartY, _nbDistEndY, rectCenterX, rectCenterX);
            }

            // ===== ROTATED: Draw horizontal offset distance line =====
            PointF distLineStart = ImageToControlPoint(centerX, DistanceLineY, imageRect);
            PointF distLineEnd = ImageToControlPoint(startX, DistanceLineY, imageRect);

            using (Pen distLinePen = new Pen(Color.ForestGreen, DistLinePenWidth))
            {
                g.DrawLine(distLinePen, distLineStart, distLineEnd);
            }

            // Draw arrowheads at both ends
            float arrowLen = 6f;
            float arrowWidth = 4f;
            float dirX = distLineEnd.X - distLineStart.X;
            float dirY = distLineEnd.Y - distLineStart.Y;
            float len = (float)Math.Sqrt(dirX * dirX + dirY * dirY);
            if (len > 0)
            {
                dirX /= len;
                dirY /= len;
                float normX = -dirY;
                float normY = dirX;

                PointF rightWing1 = new PointF(distLineStart.X + arrowLen * dirX - arrowWidth * normX,
                                          distLineStart.Y + arrowLen * dirY - arrowWidth * normY);
                PointF rightWing2 = new PointF(distLineStart.X + arrowLen * dirX + arrowWidth * normX,
                                          distLineStart.Y + arrowLen * dirY + arrowWidth * normY);

                PointF leftWing1 = new PointF(distLineEnd.X - arrowLen * dirX + arrowWidth * normX,
                                          distLineEnd.Y - arrowLen * dirY + arrowWidth * normY);
                PointF leftWing2 = new PointF(distLineEnd.X - arrowLen * dirX - arrowWidth * normX,
                                          distLineEnd.Y - arrowLen * dirY - arrowWidth * normY);

                using (Brush brush = new SolidBrush(Color.ForestGreen))
                {
                    g.FillPolygon(brush, new PointF[] { distLineStart, rightWing1, rightWing2 });
                    g.FillPolygon(brush, new PointF[] { distLineEnd, leftWing1, leftWing2 });
                }
            }

            // ===== RESTORE TRANSFORM =====
            g.Transform = originalTransform;
            originalTransform.Dispose();

            // ===== NON-ROTATED: Position textboxes and button =====
            PositionOverlayControls(imageRect, circleCenter, circleCenterNormX, circleCenterNormY,
                centerX, startX, _tcDistStartX, _tcDistStartY, _tcDistEndY,
                _ncDistStartX, _ncDistStartY, _ncDistEndY,
                _nbDistStartX, _nbDistStartY, _nbDistEndY);

            // ===== UPDATE HOTSPOT POSITIONS (except TankCenterline and RotationArrow) =====
            // Dots are drawn in rotated space, but hit-testing uses non-rotated screen coords.
            // Apply the same rotation to dot positions so hit-test matches visual positions.
            double rotDeg = (_nozzleConfig != null ? _nozzleConfig.RotationAngleDegrees : 0.0);
            double rotRadForHotspot = rotDeg * Math.PI / 180.0;
            float pivotNormX = circleCenterNormX;
            float pivotNormY = circleCenterNormY;

            UpdateHotspotPosition(NozzlePropertiesType.Connection, RotateNormX(rectCenterX, dot1Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(rectCenterX, dot1Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotRefPosition(NozzleTopReferenceType.TankCenterline, RotateNormX(dot7X, dot7Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(dot7X, dot7Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotRefPosition(NozzleTopReferenceType.NozzleCenterline, RotateNormX(dot3X, dot3Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(dot3X, dot3Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotBottomPosition(NozzleBottomReferencePoint.Top, RotateNormX(dot4X, dot4Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(dot4X, dot4Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotBottomPosition(NozzleBottomReferencePoint.Middle, RotateNormX(dot5X, dot5Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(dot5X, dot5Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotBottomPosition(NozzleBottomReferencePoint.Bottom, RotateNormX(dot6X, dot6Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(dot6X, dot6Y, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotByFlipPosition(FlipDot.Left, RotateNormX(flipLeftDotX, flipLeftDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(flipLeftDotX, flipLeftDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotByFlipPosition(FlipDot.Right, RotateNormX(flipRightDotX, flipRightDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(flipRightDotX, flipRightDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpdateHotspotByFlipPosition(FlipDot.Central, RotateNormX(flipMiddleDotX, flipMiddleDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(flipMiddleDotX, flipMiddleDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
            UpfateHotspotByNozzleLength(RotateNormX(nozzleLengthDotX, nozzleLengthDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect), RotateNormY(nozzleLengthDotX, nozzleLengthDotY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect));
        }

        /// <summary>
        /// Updates the normalized position of the hotspot matching the given <see cref="NozzlePropertiesType"/>.
        /// </summary>
        private void UpdateHotspotPosition(NozzlePropertiesType type, float x, float y)
        {
            var hs = FindHotspot(type);
            if (hs != null) { hs.X = x; hs.Y = y; }
        }

        /// <summary>
        /// Draws a vertical distance line with arrowheads at both ends and horizontal dashed leader lines.
        /// Used by tank centerline, nozzle centerline, and top/middle/bottom distance visualizations.
        /// </summary>
        /// <param name="lineX">X coordinate of the vertical distance line in normalized coords.</param>
        /// <param name="startY">Top Y endpoint of the distance line in normalized coords.</param>
        /// <param name="endY">Bottom Y endpoint of the distance line in normalized coords.</param>
        /// <param name="dashTargetStartX">X coordinate the top dashed leader line extends to (e.g. rectangle or dot center).</param>
        /// <param name="dashTargetEndX">X coordinate the bottom dashed leader line extends to.</param>
        private void DrawDistanceLine(Graphics g, Rectangle imageRect, float lineX, float startY, float endY, float dashTargetStartX, float dashTargetEndX)
        {
            PointF start = ImageToControlPoint(lineX, startY, imageRect);
            PointF end = ImageToControlPoint(lineX, endY, imageRect);

            using (Pen distPen = new Pen(Color.ForestGreen, DistLinePenWidth))
                g.DrawLine(distPen, start, end);

            // Draw arrowheads at both ends
            float arrowLen = 6f;
            float arrowWidth = 4f;
            float dirX = end.X - start.X;
            float dirY = end.Y - start.Y;
            float len = (float)Math.Sqrt(dirX * dirX + dirY * dirY);
            if (len > 0)
            {
                dirX /= len;
                dirY /= len;
                float normX = -dirY;
                float normY = dirX;

                PointF leftWing1 = new PointF(start.X + arrowLen * dirX - arrowWidth * normX,
                                          start.Y + arrowLen * dirY - arrowWidth * normY);
                PointF leftWing2 = new PointF(start.X + arrowLen * dirX + arrowWidth * normX,
                                          start.Y + arrowLen * dirY + arrowWidth * normY);

                PointF rightWing1 = new PointF(end.X - arrowLen * dirX + arrowWidth * normX,
                                          end.Y - arrowLen * dirY + arrowWidth * normY);
                PointF rightWing2 = new PointF(end.X - arrowLen * dirX - arrowWidth * normX,
                                          end.Y - arrowLen * dirY - arrowWidth * normY);

                using (Brush brush = new SolidBrush(Color.ForestGreen))
                {
                    g.FillPolygon(brush, new PointF[] { start, leftWing1, leftWing2 });
                    g.FillPolygon(brush, new PointF[] { end, rightWing1, rightWing2 });
                }
            }

            // Draw horizontal dashed lines from each arrowhead to the target X (dot or rectangle).
            // Extend 10px beyond the distance line on the outer side for visual clarity.
            float dashExtensionPx = 10f;
            float startExtDir = Math.Sign(start.X - ImageToControlPoint(dashTargetStartX, startY, imageRect).X);
            float endExtDir = Math.Sign(end.X - ImageToControlPoint(dashTargetEndX, endY, imageRect).X);
            PointF dashStartOuter = new PointF(start.X + startExtDir * dashExtensionPx, start.Y);
            PointF dashEndOuter = new PointF(end.X + endExtDir * dashExtensionPx, end.Y);
            PointF dashStartTarget = ImageToControlPoint(dashTargetStartX, startY, imageRect);
            PointF dashEndTarget = ImageToControlPoint(dashTargetEndX, endY, imageRect);
            using (Pen dashedPen = new Pen(Color.ForestGreen, 1.5f))
            {
                dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                g.DrawLine(dashedPen, dashStartOuter, dashStartTarget);
                g.DrawLine(dashedPen, dashEndOuter, dashEndTarget);
            }
        }

        /// <summary>
        /// Draws a filled and outlined circle at the given control-space center point.
        /// </summary>
        private void DrawDot(Graphics g, PointF center, Color color)
        {
            RectangleF rect = new RectangleF(
                center.X - DotRadiusPx, center.Y - DotRadiusPx,
                DotRadiusPx * 2, DotRadiusPx * 2);
            using (Brush brush = new SolidBrush(color))
            using (Pen pen = new Pen(color, 1f))
            {
                g.FillEllipse(brush, rect);
                g.DrawEllipse(pen, rect);
            }
        }

        /// <summary>
        /// Positions all overlay controls (TextBoxes and Flip button) in non-rotated space,
        /// applying rotation offset so they track their corresponding visualization anchors.
        /// </summary>
        private void PositionOverlayControls(Rectangle imageRect, PointF circleCenter,
            float circleCenterNormX, float circleCenterNormY,
            float centerX, float startX,
            float tcDistStartX, float tcDistStartY, float tcDistEndY,
            float ncDistStartX, float ncDistStartY, float ncDistEndY,
            float nbDistStartX, float nbDistStartY, float nbDistEndY)
        {
            int pbOffX = _pictureBox.Location.X;
            int pbOffY = _pictureBox.Location.Y;

            double rotRad = (_nozzleConfig != null ? _nozzleConfig.RotationAngleDegrees : 0.0) * Math.PI / 180.0;
            float pivX = circleCenterNormX;
            float pivY = circleCenterNormY;

            float rotTbAnchorX = circleCenterNormX;
            float rotTbAnchorY = circleCenterNormY - OuterCircleRadius - CenterLineTopOverflowPx / imageRect.Height;
            float rotRotTbX = RotateNormX(rotTbAnchorX, rotTbAnchorY, pivX, pivY, rotRad, imageRect);
            float rotRotTbY = RotateNormY(rotTbAnchorX, rotTbAnchorY, pivX, pivY, rotRad, imageRect);
            PointF rotTbPos = ImageToControlPoint(rotRotTbX, rotRotTbY, imageRect);
            // Push the textbox further outward at angles where it overlaps the nozzle rectangle (e.g. 270°)
            float rotTbExtraOffsetPx = 10f * (float)Math.Abs(Math.Sin(rotRad));
            float rotTbDirX = rotTbPos.X - circleCenter.X;
            float rotTbDirY = rotTbPos.Y - circleCenter.Y;
            float rotTbDirLen = (float)Math.Sqrt(rotTbDirX * rotTbDirX + rotTbDirY * rotTbDirY);
            if (rotTbDirLen > 0)
            {
                rotTbPos.X += rotTbExtraOffsetPx * (rotTbDirX / rotTbDirLen);
                rotTbPos.Y += rotTbExtraOffsetPx * (rotTbDirY / rotTbDirLen);
            }
            _rotationTextBox.Location = new Point(
                (int)(rotTbPos.X - 2) + pbOffX,
                (int)(rotTbPos.Y - _rotationTextBox.Height / 2f) + pbOffY);

            // Position the distance TextBox anchored to the midpoint of the offset distance line
            if (_flipDot != FlipDot.Central)
            {
                _offsetTextBox.Visible = true;

                float distMidNormX = (centerX + startX) / 2f;
                float distMidNormY = DistanceLineY;
                float distTbAnchorX = distMidNormX;
                float distTbAnchorY = distMidNormY - DistPerpOffsetNorm;
                float rotDistTbX = RotateNormX(distTbAnchorX, distTbAnchorY, pivX, pivY, rotRad, imageRect);
                float rotDistTbY = RotateNormY(distTbAnchorX, distTbAnchorY, pivX, pivY, rotRad, imageRect);
                PointF distTbPos = ImageToControlPoint(rotDistTbX, rotDistTbY, imageRect);
                _offsetTextBox.Location = new Point(
                    (int)(distTbPos.X - _offsetTextBox.Width / 2f) + pbOffX,
                    (int)(distTbPos.Y - _offsetTextBox.Height / 2f) + pbOffY);
            }
            else
                _offsetTextBox.Visible = false;

            // Position tank centerline distance textbox — opposite side from other visualizations
            if (_showTankCenterlineDistance)
            {
                float tcPerpOffset = _flipDot == FlipDot.Right ? -VertDistPerpOffset : VertDistPerpOffset;
                PositionTextBoxAtAnchor(_nozzleCenterlineDistanceTextBox, imageRect, pbOffX, pbOffY,
                    tcDistStartX, (tcDistStartY + tcDistEndY) / 2f,
                    tcPerpOffset,
                    pivX, pivY, rotRad);
            }
            else if (_showNozzleCenterlineDistance)
            {
                float ncPerpOffset = _flipDot == FlipDot.Right ? -VertDistPerpOffset : VertDistPerpOffset;
                PositionTextBoxAtAnchor(_nozzleCenterlineDistanceTextBox, imageRect, pbOffX, pbOffY,
                    ncDistStartX, (ncDistStartY + ncDistEndY) / 2f,
                    ncPerpOffset,
                    pivX, pivY, rotRad);
            }
            else
            {
                if (_nozzleCenterlineDistanceTextBox.Visible)
                    _nozzleCenterlineDistanceTextBox.Visible = false;
            }
           

            // Position nozzle bottom distance textbox
            if (_showNozzleTopDistance || _showNozzleBottomDistance || _showNozzleMiddleDistance)
            {
                float nbPerpOffset = _flipDot == FlipDot.Right ? VertDistPerpOffset : -VertDistPerpOffset;
                PositionTextBoxAtAnchor(_nozzleBottomDistanceTextBox, imageRect, pbOffX, pbOffY,
                    nbDistStartX, (nbDistStartY + nbDistEndY) / 2f,
                    nbPerpOffset,
                    pivX, pivY, rotRad);
            }
            else
            {
                if (_nozzleBottomDistanceTextBox.Visible)
                    _nozzleBottomDistanceTextBox.Visible = false;
            }

            if (_showNozzleLength)
            {
                float nbPerpOffset = _flipDot == FlipDot.Right ? VertDistPerpOffset : -VertDistPerpOffset;
                PositionTextBoxAtAnchor(_lengthTextBox, imageRect, pbOffX, pbOffY,
                    nbDistStartX, (nbDistStartY + nbDistEndY) / 2f,
                    nbPerpOffset,
                    pivX, pivY, rotRad);
            }
            else
            {

                if (_lengthTextBox.Visible)
                    _lengthTextBox.Visible = false;
            }
        }

        /// <summary>
        /// Positions a TextBox at a rotated anchor point offset perpendicular to a distance line.
        /// </summary>
        private void PositionTextBoxAtAnchor(TextBox tb, Rectangle imageRect, int pbOffX, int pbOffY,
            float midX, float midY, float perpOffset, float pivX, float pivY, double rotRad)
        {
            float anchorX = midX + perpOffset;
            float rotX = RotateNormX(anchorX, midY, pivX, pivY, rotRad, imageRect);
            float rotY = RotateNormY(anchorX, midY, pivX, pivY, rotRad, imageRect);
            PointF pos = ImageToControlPoint(rotX, rotY, imageRect);
            Point newLocation = new Point(
                (int)(pos.X - tb.Width / 2f) + pbOffX,
                (int)(pos.Y - tb.Height / 2f) + pbOffY);
            if (tb.Location != newLocation)
                tb.Location = newLocation;
            if (!tb.Visible)
                tb.Visible = true;
        }

        /// <summary>
        /// Returns the flipped or normal offset from a base X coordinate.
        /// When flipped, adds the offset; when normal, subtracts it.
        /// </summary>
        private float FlipOffset(float baseX, float offset)
        {
            return _flipDot == FlipDot.Right ? (baseX + offset) : (baseX - offset);
        }

        /// <summary>
        /// Updates the normalized position of the hotspot matching the given <see cref="NozzleTopReferenceType"/>.
        /// </summary>
        private void UpdateHotspotRefPosition(NozzleTopReferenceType type, float x, float y)
        {
            var hs = FindHotspotByRef(type);
            if (hs != null) { hs.X = x; hs.Y = y; }
        }

        /// <summary>
        /// Updates the normalized position of the hotspot matching the given <see cref="FlipDot"/> value.
        /// </summary>
        private void UpdateHotspotByFlipPosition(FlipDot flipDot, float x, float y)
        {
            var hs = FindHotspotByFlipDot(flipDot);
            if (hs != null) { hs.X = x; hs.Y = y; }
        }

        /// <summary>
        /// Updates the normalized position of the hotspot matching the given <see cref="NozzleBottomReferencePoint"/>.
        /// </summary>
        private void UpdateHotspotBottomPosition(NozzleBottomReferencePoint type, float x, float y)
        {
            var hs = FindHotspotByBottom(type);
            if (hs != null) { hs.X = x; hs.Y = y; }
        }

        /// <summary>
        /// Updates the normalized position of the nozzle length hotspot.
        /// </summary>
        private void UpfateHotspotByNozzleLength(float x, float y)
        {
            var hs = FindHotspotByNozzleLength();
            if (hs != null) { hs.X = x; hs.Y = y; }
        }

        /// <summary>
        /// Rotates a point (in normalized coords) around a pivot by the given angle,
        /// accounting for the aspect ratio difference between X and Y in pixel space.
        /// Returns the rotated X in normalized coords.
        /// </summary>
        private float RotateNormX(float normX, float normY, float pivotNormX, float pivotNormY, double angleRad, Rectangle imageRect)
        {
            float dxPx = (normX - pivotNormX) * imageRect.Width;
            float dyPx = (normY - pivotNormY) * imageRect.Height;
            double cos = Math.Cos(angleRad);
            double sin = Math.Sin(angleRad);
            double rxPx = dxPx * cos - dyPx * sin;
            return pivotNormX + (float)(rxPx / imageRect.Width);
        }

        /// <summary>
        /// Rotates a point (in normalized coords) around a pivot by the given angle,
        /// accounting for the aspect ratio difference between X and Y in pixel space.
        /// Returns the rotated Y in normalized coords.
        /// </summary>
        private float RotateNormY(float normX, float normY, float pivotNormX, float pivotNormY, double angleRad, Rectangle imageRect)
        {
            float dxPx = (normX - pivotNormX) * imageRect.Width;
            float dyPx = (normY - pivotNormY) * imageRect.Height;
            double cos = Math.Cos(angleRad);
            double sin = Math.Sin(angleRad);
            double ryPx = dxPx * sin + dyPx * cos;
            return pivotNormY + (float)(ryPx / imageRect.Height);
        }

        /// <summary>
        /// Pushes all visible textbox values into the model.
        /// Called before switching reference points so the user's typed value is committed.
        /// </summary>
        private void CommitAllTextBoxes()
        {
            CommitTextBoxToModel(_offsetTextBox);
            CommitTextBoxToModel(_nozzleCenterlineDistanceTextBox);
            CommitTextBoxToModel(_nozzleBottomDistanceTextBox);
            CommitTextBoxToModel(_rotationTextBox);
            CommitTextBoxToModel(_lengthTextBox);
        }

        /// <summary>
        /// Parses the textbox text (mm) and writes to the corresponding model property (meters).
        /// </summary>
        
        private void CommitTextBoxToModel(TextBox tb)
        {
            if (tb == null || !tb.Visible || _nozzleConfig == null) return;

            var txt = (tb.Text ?? string.Empty).Replace("°", "").Trim();
            if (string.IsNullOrEmpty(txt)) return;
            txt = txt.Replace(",", ".");

            if (tb == _rotationTextBox)
            {
                if (!double.TryParse(txt, NumberStyles.Float,
                    CultureInfo.CurrentCulture, out double angle))
                    return;

                // Normalize angle to 0-360 degrees (using modulo operator)
                angle = (angle % 360 + 360) % 360;

                _nozzleConfig.RotationAngleDegrees = angle;

                return;
            }

            if (!double.TryParse(txt, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture, out double mm))
                return;

            double meters = mm / 1000.0;

            if (tb == _offsetTextBox)
                _nozzleConfig.OffsetMeters = meters;
            else if (tb == _nozzleCenterlineDistanceTextBox)
                _nozzleConfig.DistanceFromTopReferenceMeters = meters;
            else if (tb == _nozzleBottomDistanceTextBox)
                _nozzleConfig.DistanceFromBottomReferenceMeters = meters;
            else if (tb == _lengthTextBox)
                _nozzleConfig.NozzleLength = meters;
        }

        /// <summary>
        /// Reads the model property (meters) and writes to the textbox (mm).
        /// </summary>
        
        private void UpdateTextBoxFromModel(TextBox tb)
        {
            if (tb == null || _nozzleConfig == null) return;

            double meters = 0.0;
            double angle = 0.0;
            if (tb == _offsetTextBox)
                meters = _nozzleConfig.OffsetMeters;
            else if (tb == _nozzleCenterlineDistanceTextBox)
                meters = _nozzleConfig.DistanceFromTopReferenceMeters;
            else if (tb == _nozzleBottomDistanceTextBox)
                meters = _nozzleConfig.DistanceFromBottomReferenceMeters;
            else if (tb == _rotationTextBox)
                angle = _nozzleConfig.RotationAngleDegrees;
            else if (tb == _lengthTextBox)
                meters = _nozzleConfig.NozzleLength;

            if ((tb == _nozzleCenterlineDistanceTextBox || tb == _nozzleBottomDistanceTextBox || tb == _lengthTextBox)
                && meters == 0.0)
            {
                tb.Text = string.Empty;
            }
            else if (tb == _rotationTextBox)
            {
                tb.Text = $"{angle.ToString(CultureInfo.CurrentCulture)}°";
            }
            else
            {
                tb.Text = (meters * 1000.0).ToString(CultureInfo.CurrentCulture);
            }
        }

        /// <summary>
        /// Activates the tank centerline distance visualization, hides the nozzle centerline one,
        /// shows the distance TextBox, and raises the DistanceChanged event.
        /// </summary>
        public void ToggleTankCenterlineDistanceVisualization()
        {
            CommitAllTextBoxes();
            _showNozzleCenterlineDistance = false;
            _showTankCenterlineDistance = true;
            _pictureBox.Invalidate();

            DistanceChanged?.Invoke(this, new DistanceChangedEventArgs
            {
                TopReferenceType = NozzleTopReferenceType.TankCenterline
            });

            if (_nozzleCenterlineDistanceTextBox != null)
            {
                _nozzleCenterlineDistanceTextBox.Visible = true;
                UpdateTextBoxFromModel(_nozzleCenterlineDistanceTextBox);
            }
        }

        /// <summary>
        /// Activates the nozzle centerline distance visualization, hides the tank centerline one,
        /// shows the distance TextBox, and raises the DistanceChanged event.
        /// </summary>
        public void ToggleNozzleCenterlineDistanceVisualization()
        {
            CommitAllTextBoxes();
            _showTankCenterlineDistance = false;
            _showNozzleCenterlineDistance = true;
            _pictureBox.Invalidate();

            DistanceChanged?.Invoke(this, new DistanceChangedEventArgs
            {
                TopReferenceType = NozzleTopReferenceType.NozzleCenterline
            });

            if (_nozzleCenterlineDistanceTextBox != null)
            {
                _nozzleCenterlineDistanceTextBox.Visible = true;
                UpdateTextBoxFromModel(_nozzleCenterlineDistanceTextBox);
            }
        }

        /// <summary>
        /// Activates the nozzle top (BottomReferencePoint.Top) distance visualization with a short rectangle,
        /// hides middle and bottom visualizations, and raises the DistanceChanged event.
        /// On consecutive clicks, toggles the diagonal cut on the nozzle rectangle.
        /// The diagonal cut state is preserved when switching to other visualizations and restored on return.
        /// </summary>
        public void ToggleNozzleTopDistanceVisualization()
        {
            // Toggle diagonal cut if top dot is clicked consecutively
            if (_showNozzleTopDistance && _lastActiveNozzleVisualization == VisualizationTop)
            {
                _isDiagonalCut = !_isDiagonalCut;
            }

            CommitAllTextBoxes();
            _showNozzleLength = false;
            _showNozzleMiddleDistance = false;
            _showNozzleBottomDistance = false;
            _showNozzleTopDistance = true;
            _lastActiveNozzleVisualization = VisualizationTop;
            _pictureBox.Invalidate();

            DistanceChanged?.Invoke(this, new DistanceChangedEventArgs
            {
                BottomReferencePoint = NozzleBottomReferencePoint.Top
            });

            if (_nozzleBottomDistanceTextBox != null)
            {
                _nozzleBottomDistanceTextBox.Visible = true;
                UpdateTextBoxFromModel(_nozzleBottomDistanceTextBox);
            }
        }

        /// <summary>
        /// Activates the nozzle bottom (BottomReferencePoint.Bottom) distance visualization with a long rectangle,
        /// hides top and middle visualizations, and raises the DistanceChanged event.
        /// </summary>
        public void ToggleNozzleBottomDistanceVisualization()
        {
            CommitAllTextBoxes();
            _showNozzleLength = false;
            _showNozzleMiddleDistance = false;
            _showNozzleTopDistance = false;
            _showNozzleBottomDistance = true;
            _lastActiveNozzleVisualization = VisualizationBottom;
            _pictureBox.Invalidate();

            DistanceChanged?.Invoke(this, new DistanceChangedEventArgs
            {
                BottomReferencePoint = NozzleBottomReferencePoint.Bottom
            });

            if (_nozzleBottomDistanceTextBox != null)
            {
                _nozzleBottomDistanceTextBox.Visible = true;
                UpdateTextBoxFromModel(_nozzleBottomDistanceTextBox);
            }
        }

        /// <summary>
        /// Activates the nozzle middle (BottomReferencePoint.Middle) distance visualization.
        /// Toggles between short and long rectangle on consecutive clicks.
        /// Hides top and bottom visualizations, and raises the DistanceChanged event.
        /// </summary>
        public void ToggleNozzleMiddleDistanceVisualization()
        {
            CommitAllTextBoxes();
            _showNozzleLength = false;
            _showNozzleTopDistance = false;
            _showNozzleBottomDistance = false;
            _showNozzleMiddleDistance = true;
            
            // Only toggle rectangle size if middle dot is clicked consecutively
            if (_lastActiveNozzleVisualization == VisualizationMiddle)
            {
                _isNozzleMiddleRectangleLong = !_isNozzleMiddleRectangleLong;
            }
            
            _lastActiveNozzleVisualization = VisualizationMiddle;
            _pictureBox.Invalidate();

            DistanceChanged?.Invoke(this, new DistanceChangedEventArgs
            {
                BottomReferencePoint = NozzleBottomReferencePoint.Middle
            });

            if (_nozzleBottomDistanceTextBox != null)
            {
                _nozzleBottomDistanceTextBox.Visible = true;
                UpdateTextBoxFromModel(_nozzleBottomDistanceTextBox);
            }
        }

        /// <summary>
        /// Refreshes the rotation TextBox from the model and repaints the visualization
        /// to reflect the current rotation angle.
        /// </summary>
        public void ToggleRotationVisualization()
        {
            UpdateTextBoxFromModel(_rotationTextBox);
            _pictureBox.Invalidate();
        }

        /// <summary>
        /// Activates the nozzle length visualization, hides top/middle/bottom distance visualizations,
        /// shows the length TextBox, and repaints.
        /// </summary>
        public void ToggleNozzleLengthVisualization()
        {
            CommitAllTextBoxes();
            _showNozzleLength = true;
            _showNozzleTopDistance = false;
            _showNozzleBottomDistance = false;
            _showNozzleMiddleDistance = false;

            _pictureBox.Invalidate();

            if (_lengthTextBox != null)
            {
                _lengthTextBox.Visible = true;
                UpdateTextBoxFromModel(_lengthTextBox);
            }
        }

        /// <summary>
        /// Enables or disables all overlay visualizations (textboxes, flip button, painted overlays, hotspot clicks).
        /// When disabled, the control appears blank; when re-enabled, it redraws.
        /// </summary>
        public void SetVisualizationsEnabled(bool enabled)
        {
            _visualizationsEnabled = enabled;

            _offsetTextBox.Visible = enabled;
            _rotationTextBox.Visible = enabled;
            _nozzleCenterlineDistanceTextBox.Visible = enabled && (_showTankCenterlineDistance || _showNozzleCenterlineDistance);
            _nozzleBottomDistanceTextBox.Visible = enabled && (_showNozzleTopDistance || _showNozzleBottomDistance || _showNozzleMiddleDistance);
            _lengthTextBox.Visible = enabled && _showNozzleLength;
        }
    }
}
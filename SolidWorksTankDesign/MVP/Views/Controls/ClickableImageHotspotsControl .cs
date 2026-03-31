using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SolidWorksTankDesign.MVP.Enums;
using System.Globalization;
using System.Linq;
using SolidWorksTankDesign.TankSiteConfigurations;

namespace SolidWorksTankDesign.MVP.Views.Controls
{
    public class ClickableImageHotspotsControl : UserControl
    {
        private PictureBox _pictureBox;
        private List<Hotspot> _hotspots;
        private TextBox _offsetTextBox;
        private TextBox _nozzleCenterlineDistanceTextBox;
        private TextBox _nozzleBottomDistanceTextBox;
        private TextBox _rotationTextBox;
        private bool _showTankCenterlineDistance = false;
        private bool _showNozzleCenterlineDistance = false;
        private bool _showNozzleTopDistance = false;
        private bool _showNozzleMiddleDistance = false;
        private bool _showNozzleBottomDistance = false;
        private bool _isNozzleMiddleRectangleLong = false;
        private string _lastActiveNozzleVisualization = null;
        private bool _isFlipped = false;

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
        private const float CenterLineOverflowPx = 15f;
        private const float TankCenterlineDotOffsetPx = 15f;
        private const float DistVisualizationOffsetNorm = 0.15f;
        private const float DistVisualizationHeightPx = 40f;
        private const float DistanceLineY = 0.8f;
        private const float CapWidth = 0.04f;
        private const float DistLinePenWidth = 1.7f;
        private const float FlipArrowPenWidth = 2.5f;
        private const float VertDistPerpOffset = 0.1f;
        private const float DistPerpOffsetNorm = 0.06f;
        private const float ArrowArcRadiusPx = 15f;

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

        public ClickableImageHotspotsControl()
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
            _pictureBox.Paint += PictureBox_Paint;
            Controls.Add(_pictureBox);

            // Add distance input TextBox
            _offsetTextBox = new TextBox
            {
                Name = "DistanceTextBox",
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White
            };

            _offsetTextBox.LostFocus += _distanceTextBox_LostFocus;
            _offsetTextBox.MouseClick += _distanceTextBox_MouseClick;

            // Position will be set dynamically in Paint event
            Controls.Add(_offsetTextBox);
            _offsetTextBox.BringToFront();  // ensure it's on top of PictureBox

            _nozzleCenterlineDistanceTextBox = new TextBox
            {
                Name = "CenterlineDistanceTextBox",
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
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
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
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
                Width = 30,
                Height = 15,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = HorizontalAlignment.Center,
                BackColor = Color.White,
                Visible = true
            };
            _rotationTextBox.LostFocus += _rotationTextBox_LostFocus;
            _rotationTextBox.MouseClick += _rotationTextBox_MouseClick;
            _rotationTextBox.PreviewKeyDown += _rotationTextBox_PreviewKeyDown;
            _rotationTextBox.KeyDown += _rotationTextBox_KeyDown;

            // Position will be set dynamically in Paint event
            Controls.Add(_rotationTextBox);
            _rotationTextBox.BringToFront();
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
        public void ToggleFlip()
        {
            _isFlipped = !_isFlipped;
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
        /// Finds the first hotspot matching the given <see cref="NozzleBottomReferencePoint"/>, or null if not found.
        /// </summary>
        private Hotspot FindHotspotByBottom(NozzleBottomReferencePoint type)
        {
            foreach (var hs in _hotspots)
                if (hs.BottomReferencePoint == type) return hs;
            return null;
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

            // Hit-test against dynamically computed dot positions (in sync with Paint)
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
            else if (hotspot.IsFlipArrow)
            {
                ToggleFlip();
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
        private Rectangle GetImageRect()
        {
            if (_pictureBox.Image == null)
            {
                // Center a fixed-size drawing area within the PictureBox so that
                // visualizations render at the same size as before the padding was removed.
                int x = (_pictureBox.ClientSize.Width - DrawingAreaSize) / 2;
                int y = (_pictureBox.ClientSize.Height - DrawingAreaSize) / 2;
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
            float startX = _isFlipped ? (centerX + (centerX - NozzleLineBaseX)) : NozzleLineBaseX;
            float startY = NozzleLineStartY;
            float endX = startX;
            float endY = NozzleLineEndY;

            // Convert to control pixel coordinates
            PointF start = ImageToControlPoint(startX, startY, imageRect);
            PointF end = ImageToControlPoint(endX, endY, imageRect);

            // Add a rectangle at the midpoint of the line
            // Determine rectangle size based on which reference point is active
            float rectCenterX = _isFlipped ? (centerX + (centerX - NozzleLineBaseX)) : NozzleLineBaseX;
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
                endY = circleCenterNormY + (float)(t / imageRect.Height);
            }

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
                PointF vTop = new PointF(circleCenter.X, circleCenter.Y - outerCircleRadiusPx - CenterLineOverflowPx);
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

            // Draw the rectangle
            using (Pen rectPen = new Pen(Color.Gray, 2f))
            {
                g.DrawRectangle(rectPen, Rectangle.Round(rect));
            }

            // ===== DOTS (drawn on top of circles, still rotated) =====

            // --- Dot 1: Connection (middle of the top border of the rectangle) ---
            float dot1Y = rectCenterY - rectHeight / 2;
            DrawDot(g, ImageToControlPoint(rectCenterX, dot1Y, imageRect), Color.Black);

            // --- Dot 3: NozzleCenterline - inward from line start ---
            float dot3InsetNorm = Dot3InsetPx / imageRect.Height;
            float dot3Y = startY + dot3InsetNorm;
            float dot3X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot3X, dot3Y, imageRect), Color.FromArgb(80, 135, 138));

            // --- Dot 4: BottomReferencePoint.Top - inward from line start ---
            float dotInsetNorm = DotInsetPx / imageRect.Height;
            float dot4OrigY = startY + dotInsetNorm + DotInsetExtra;
            float dot4Y = dot4OrigY;
            float dot4X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot4X, dot4Y, imageRect), Color.Purple);

            // --- Dot 5: BottomReferencePoint.Middle ---
            float dot5OrigY = circleCenterNormY;
            float dot5X = rectCenterX;
            float dot5Y = dot5OrigY;
            DrawDot(g, ImageToControlPoint(dot5X, dot5Y, imageRect), Color.Purple);

            // --- Dot 6: BottomReferencePoint.Bottom - inward from line end ---
            float dot6Y = endY - dotInsetNorm - DotInsetExtra;
            float dot6OrigY = dot6Y;
            float dot6X = rectCenterX;
            DrawDot(g, ImageToControlPoint(dot6X, dot6Y, imageRect), Color.Purple);

            // ===== ROTATED: Distance visualization lines (drawn in rotated space) =====

            // --- Dot 7: TankCenterline - where vertical center line meets outer circle border (top) ---
            float dot7X = circleCenterNormX;
            float dot7Y = circleCenterNormY - outerCircleRadiusPx / imageRect.Height;
            DrawDot(g, ImageToControlPoint(dot7X, dot7Y, imageRect), Color.FromArgb(80, 135, 138));

            // ===== ROTATED: Rotation arrow =====

            PointF rotationArrowLocation = DrawRotationArrow(g, circleCenter);


            // Compute dotX/dotY for tank centerline and nozzle centerline visualization sections
            // Position tank centerline dot 5px from the outer circle's vertical center line
            float tankCenterlineDotOffsetNorm = TankCenterlineDotOffsetPx / imageRect.Width;
            float dotX = _isFlipped
                ? circleCenterNormX - tankCenterlineDotOffsetNorm
                : circleCenterNormX + tankCenterlineDotOffsetNorm;
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
            float baseDot4OrigY = baseStartY + baseDotInsetNormVal;
            float baseDot5OrigY = circleCenterNormY;
            float baseDot6OrigY = baseEndY - baseDotInsetNormVal;

            float baseDot3InsetNorm = Dot3InsetPx / imageRect.Height;
            float baseDot3Y = baseStartY + baseDot3InsetNorm;

            // Store pre-rotation coords for nozzle centerline visualization using base values
            if (_showNozzleCenterlineDistance)
            {
                _ncDistStartX = FlipOffset(rectCenterX, DistVisualizationOffsetNorm);
                _ncDistStartY = baseDot3Y;
                _ncDistEndY = baseDot3Y - DistVisualizationHeightPx / imageRect.Height;
            }

            // Store pre-rotation coords for nozzle top/bottom/middle visualization using base values
            if (_showNozzleTopDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, DistVisualizationOffsetNorm);
                _nbDistStartY = baseDot4OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }
            else if (_showNozzleBottomDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, DistVisualizationOffsetNorm);
                _nbDistStartY = baseDot6OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }
            else if (_showNozzleMiddleDistance)
            {
                _nbDistStartX = FlipOffset(rectCenterX, DistVisualizationOffsetNorm);
                _nbDistStartY = baseDot5OrigY;
                _nbDistEndY = baseRectCenterY + baseRectHeight / 2;
            }

            // ===== ROTATED: Draw distance visualizations =====

            float flipArrowSpan = 0.1f;
            float flipArrowY = circleCenterNormY + OuterCircleRadius + CenterLineOverflowPx / imageRect.Height + 0.02f;
            float flipArrowStartX = circleCenterNormX - flipArrowSpan;
            float flipArrowEndX = circleCenterNormX + flipArrowSpan;
            DrawFlipArrow(g, imageRect, flipArrowStartX, flipArrowEndX, flipArrowY);

            // Draw tank centerline distance visualization (rotated)
            if (_showTankCenterlineDistance)
            {
                float tcLineX = dotX;
                float tcStartY = baseRectCenterY - baseRectHeight / 2;
                float tcLineDxPx2 = (tcLineX - circleCenterNormX) * imageRect.Width;
                float tcLineDyPx2 = (float)Math.Sqrt(outerRadiusPx * outerRadiusPx - tcLineDxPx2 * tcLineDxPx2);
                float tcEndY = circleCenterNormY - tcLineDyPx2 / imageRect.Height;

                _tcDistStartX = tcLineX; _tcDistStartY = tcStartY; _tcDistEndY = tcEndY;

                DrawDistanceLine(g, imageRect, tcLineX, tcStartY, tcEndY, _isFlipped, true);
            }

            // Draw nozzle centerline distance visualization (rotated)
            if (_showNozzleCenterlineDistance)
            {
                DrawDistanceLine(g, imageRect, _ncDistStartX, _ncDistStartY, _ncDistEndY, !_isFlipped);
            }

            // Draw nozzle top/bottom/middle distance visualization (rotated)
            if (_showNozzleTopDistance || _showNozzleBottomDistance || _showNozzleMiddleDistance)
            {
                DrawDistanceLine(g, imageRect, _nbDistStartX, _nbDistStartY, _nbDistEndY, !_isFlipped);
            }

            // ===== ROTATED: Draw horizontal offset distance line =====
            PointF distLineStart = ImageToControlPoint(centerX, DistanceLineY, imageRect);
            PointF distLineEnd = ImageToControlPoint(startX, DistanceLineY, imageRect);

            using (Pen distLinePen = new Pen(Color.ForestGreen, DistLinePenWidth))
            {
                g.DrawLine(distLinePen, distLineStart, distLineEnd);
            }

            PointF capLeftTop = ImageToControlPoint(centerX, DistanceLineY - CapWidth / 2, imageRect);
            PointF capLeftBottom = ImageToControlPoint(centerX, DistanceLineY + CapWidth / 2, imageRect);
            using (Pen capPen = new Pen(Color.ForestGreen, DistLinePenWidth))
            {
                g.DrawLine(capPen, capLeftTop, capLeftBottom);
            }

            PointF capRightTop = ImageToControlPoint(startX, DistanceLineY - CapWidth / 2, imageRect);
            PointF capRightBottom = ImageToControlPoint(startX, DistanceLineY + CapWidth / 2, imageRect);
            using (Pen capPen2 = new Pen(Color.ForestGreen, DistLinePenWidth))
            {
                g.DrawLine(capPen2, capRightTop, capRightBottom);
            }

            // ===== RESTORE TRANSFORM =====
            g.Transform = originalTransform;
            originalTransform.Dispose();

            // ===== NON-ROTATED: Position textboxes and button =====
            PositionOverlayControls(imageRect, circleCenter, circleCenterNormX, circleCenterNormY,
                centerX, startX, _tcDistStartX, _tcDistStartY, _tcDistEndY,
                _ncDistStartX, _ncDistStartY, _ncDistEndY,
                _nbDistStartX, _nbDistStartY, _nbDistEndY,
                rotationArrowLocation.X, rotationArrowLocation.Y);

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

            // Update rotation arrow hotspot to match the actual circle center position (non-rotated)
            var rotArrowHs = _hotspots.FirstOrDefault(h => h.IsRotationArrow);
            if (rotArrowHs != null)
            {
                rotArrowHs.X = circleCenterNormX;
                rotArrowHs.Y = circleCenterNormY;
            }

            // Update flip arrow hotspot to rotated screen position
            var flipArrowHs = _hotspots.FirstOrDefault(h => h.IsFlipArrow);
            if (flipArrowHs != null)
            {
                float flipMidX = (flipArrowStartX + flipArrowEndX) / 2f;
                flipArrowHs.X = RotateNormX(flipMidX, flipArrowY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect);
                flipArrowHs.Y = RotateNormY(flipMidX, flipArrowY, pivotNormX, pivotNormY, rotRadForHotspot, imageRect);
            }
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
        /// <param name="dashToRight">When true, dashed lines extend to the right; when false, to the left.</param>
        /// <param name="shortenBottomDash">When true, the bottom dashed line is drawn at half length.</param>
        private void DrawDistanceLine(Graphics g, Rectangle imageRect, float lineX, float startY, float endY, bool dashToRight = true, bool shortenBottomDash = false)
        {
            PointF start = ImageToControlPoint(lineX, startY, imageRect);
            PointF end = ImageToControlPoint(lineX, endY, imageRect);

            using (Pen distPen = new Pen(Color.ForestGreen, DistLinePenWidth))
                g.DrawLine(distPen, start, end);

            // Draw arrowhead at the end (right side when not flipped, left side when flipped)
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

            // Draw horizontal dashed lines from each arrowhead toward the nozzle line
            float dashOffset = dashToRight ? 30f : -30f;
            float bottomDashOffset = shortenBottomDash ? dashOffset / 2f : dashOffset;
            using (Pen dashedPen = new Pen(Color.ForestGreen, 1.5f))
            {
                dashedPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                g.DrawLine(dashedPen, start, new PointF(start.X + dashOffset, start.Y));
                g.DrawLine(dashedPen, end, new PointF(end.X + bottomDashOffset, end.Y));
            }
        }

        /// <summary>
        /// Draws a horizontal flip arrow with arrowheads at both ends.
        /// Also updates the flip arrow hotspot position to the center of the arrow line.
        /// </summary>
        private void DrawFlipArrow(Graphics g, Rectangle imageRect, float startX, float endX, float lineY)
        {
            PointF start = ImageToControlPoint(startX, lineY, imageRect);
            PointF end = ImageToControlPoint(endX, lineY, imageRect);

            using (Pen distPen = new Pen(Color.RoyalBlue, FlipArrowPenWidth))
                g.DrawLine(distPen, start, end);

            // Draw arrowhead at the end (right side when not flipped, left side when flipped)
            float arrowLen = 8f;
            float arrowWidth = 6f;
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

                using (Brush brush = new SolidBrush(Color.RoyalBlue))
                {
                    g.FillPolygon(brush, new PointF[] { start, leftWing1, leftWing2 });
                    g.FillPolygon(brush, new PointF[] { end, rightWing1, rightWing2 });
                }
                    
            }

            // Update the flip arrow hotspot position to the center of the arrow line
            var flipHs = _hotspots.FirstOrDefault(h => h.IsFlipArrow);
            if (flipHs != null)
            {
                flipHs.X = (startX + endX) / 2f;
                flipHs.Y = lineY;
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
        /// Draws a semi-circular rotation arrow indicator at the given center point.
        /// </summary>
        private PointF DrawRotationArrow(Graphics g, PointF circleCenter)
        {
            RectangleF arcRect = new RectangleF(
                circleCenter.X - ArrowArcRadiusPx,
                circleCenter.Y - ArrowArcRadiusPx,
                ArrowArcRadiusPx * 2,
                ArrowArcRadiusPx * 2);

            float arcStartAngle = -30f;
            float arcSweepAngle = 300f;

            using (Pen arcPen = new Pen(Color.Black, 1.5f))
            {
                g.DrawArc(arcPen, arcRect, arcStartAngle, arcSweepAngle);
            }

            double endAngleRad = (arcStartAngle + arcSweepAngle) * Math.PI / 180.0;
            float arrowTipX = circleCenter.X + ArrowArcRadiusPx * (float)Math.Cos(endAngleRad);
            float arrowTipY = circleCenter.Y + 1.5f + ArrowArcRadiusPx * (float)Math.Sin(endAngleRad);

            float tangentX = -(float)Math.Sin(endAngleRad);
            float tangentY = (float)Math.Cos(endAngleRad);

            float arrowLen = 7f;
            float arrowWidth = 3.5f;

            float normalX = -tangentY;
            float normalY = tangentX;

            PointF wing1 = new PointF(
                arrowTipX - arrowLen * tangentX + arrowWidth * normalX,
                arrowTipY - arrowLen * tangentY + arrowWidth * normalY);
            PointF wing2 = new PointF(
                arrowTipX - arrowLen * tangentX - arrowWidth * normalX,
                arrowTipY - arrowLen * tangentY - arrowWidth * normalY);

            PointF tip = new PointF(arrowTipX, arrowTipY);
            using (Brush arrowBrush = new SolidBrush(Color.Black))
            {
                g.FillPolygon(arrowBrush, new PointF[] { tip, wing1, wing2 });
            }

            return tip;
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
            float nbDistStartX, float nbDistStartY, float nbDistEndY,
            float rotationArrowX, float rotationArrowY)
        {
            int pbOffX = _pictureBox.Location.X;
            int pbOffY = _pictureBox.Location.Y;

            double rotRad = (_nozzleConfig != null ? _nozzleConfig.RotationAngleDegrees : 0.0) * Math.PI / 180.0;
            float pivX = circleCenterNormX;
            float pivY = circleCenterNormY;

            // In PositionOverlayControls, replace the rotation textbox positioning with:
            float rotTbOffsetX = _isFlipped ? -0.075f : 0.075f;
            float rotTbAnchorX = circleCenterNormX + rotTbOffsetX;
            float rotTbAnchorY = circleCenterNormY - 0.1f;
            float rotRotTbX = RotateNormX(rotTbAnchorX, rotTbAnchorY, pivX, pivY, rotRad, imageRect);
            float rotRotTbY = RotateNormY(rotTbAnchorX, rotTbAnchorY, pivX, pivY, rotRad, imageRect);
            PointF rotTbPos = ImageToControlPoint(rotRotTbX, rotRotTbY, imageRect);
            _rotationTextBox.Location = new Point(
                (int)(rotTbPos.X - _rotationTextBox.Width / 2f) + pbOffX,
                (int)(rotTbPos.Y - _rotationTextBox.Height / 2f) + pbOffY);

            // Position the distance TextBox anchored to the midpoint of the offset distance line
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

            // Position the Flip button anchored to the opposite side of the offset distance line
            float flipTbAnchorY = distMidNormY + DistPerpOffsetNorm + VertDistPerpOffset;
            float rotFlipX = RotateNormX(distMidNormX, flipTbAnchorY, pivX, pivY, rotRad, imageRect);
            float rotFlipY = RotateNormY(distMidNormX, flipTbAnchorY, pivX, pivY, rotRad, imageRect);
            PointF flipPos = ImageToControlPoint(rotFlipX, rotFlipY, imageRect);
            //_flipButton.Location = new Point(
            //    (int)(flipPos.X - _flipButton.Width / 2f) + pbOffX,
            //    (int)(flipPos.Y - _flipButton.Height / 2f) + pbOffY);

            // Position tank/nozzle centerline distance textbox
            if (_showTankCenterlineDistance)
            {
                PositionTextBoxAtAnchor(_nozzleCenterlineDistanceTextBox, imageRect, pbOffX, pbOffY,
                    tcDistStartX, (tcDistStartY + tcDistEndY) / 2f,
                    _isFlipped ? -VertDistPerpOffset : VertDistPerpOffset,
                    pivX, pivY, rotRad);
            }
            else if (_showNozzleCenterlineDistance)
            {
                PositionTextBoxAtAnchor(_nozzleCenterlineDistanceTextBox, imageRect, pbOffX, pbOffY,
                    ncDistStartX, (ncDistStartY + ncDistEndY) / 2f,
                    _isFlipped ? VertDistPerpOffset : -VertDistPerpOffset,
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
                PositionTextBoxAtAnchor(_nozzleBottomDistanceTextBox, imageRect, pbOffX, pbOffY,
                    nbDistStartX, (nbDistStartY + nbDistEndY) / 2f,
                    _isFlipped ? VertDistPerpOffset : -VertDistPerpOffset,
                    pivX, pivY, rotRad);
            }
            else
            {
                if (_nozzleBottomDistanceTextBox.Visible)
                    _nozzleBottomDistanceTextBox.Visible = false;
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
            return _isFlipped ? (baseX + offset) : (baseX - offset);
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
        /// Updates the normalized position of the hotspot matching the given <see cref="NozzleBottomReferencePoint"/>.
        /// </summary>
        private void UpdateHotspotBottomPosition(NozzleBottomReferencePoint type, float x, float y)
        {
            var hs = FindHotspotByBottom(type);
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
        }

        /// <summary>
        /// Parses the textbox text (mm) and writes to the corresponding model property (meters).
        /// </summary>
        
        private void CommitTextBoxToModel(TextBox tb)
        {
            if (tb == null || !tb.Visible || _nozzleConfig == null) return;

            var txt = (tb.Text ?? string.Empty).Trim();
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

            if ((tb == _nozzleCenterlineDistanceTextBox || tb == _nozzleBottomDistanceTextBox)
                && meters == 0.0)
            {
                tb.Text = string.Empty;
            }
            else if (tb == _rotationTextBox)
            {
                tb.Text = angle.ToString(CultureInfo.CurrentCulture);
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
        /// </summary>
        public void ToggleNozzleTopDistanceVisualization()
        {
            CommitAllTextBoxes();
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

            _pictureBox.Invalidate();
        }
    }
}
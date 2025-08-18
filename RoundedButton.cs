using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace count_dead_sign
{
    public class RoundedButton : Button
    {
        private Color buttonBackColor = Color.LightGray;
        private Color textColor = Color.White;
        private int cornerRadius = 10;
        private bool isHover = false;
        private int shadowOffset = 4;
        private int depthOffset = 2;
        private bool isPressed = false;

        private int shadowBlur = 8;
        private int hoverShadowExpand = 2;
        private Color shadowColor = Color.FromArgb(40, 0, 0, 0);

        private int borderSize = 1;
        private Color borderColor = Color.Black;
        private Color hoverBorderColor = Color.Red;

        // 사용자 정의 속성들
        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color ButtonBackColor { get => buttonBackColor; set { buttonBackColor = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color TextColor { get => textColor; set { textColor = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int CornerRadius { get => cornerRadius; set { cornerRadius = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int ShadowOffset { get => shadowOffset; set { shadowOffset = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int DepthOffset { get => depthOffset; set { depthOffset = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int ShadowBlur { get => shadowBlur; set { shadowBlur = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int HoverShadowExpand { get => hoverShadowExpand; set { hoverShadowExpand = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color ShadowColor { get => shadowColor; set { shadowColor = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int BorderSize { get => borderSize; set { borderSize = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color BorderColor { get => borderColor; set { borderColor = value; Invalidate(); } }

        [Category("Appearance")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color HoverBorderColor { get => hoverBorderColor; set { hoverBorderColor = value; Invalidate(); } }

        public RoundedButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            BackColor = Color.Transparent;
            ForeColor = Color.White;

            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.DoubleBuffer, true);
            SetStyle(ControlStyles.ResizeRedraw, true);
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);

            MouseEnter += (s, e) => { isHover = true; Invalidate(); };
            MouseLeave += (s, e) => { isHover = false; Invalidate(); };
        }

        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            base.OnMouseDown(mevent);
            isPressed = true;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            base.OnMouseUp(mevent);
            isPressed = false;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            pevent.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            pevent.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            pevent.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // 배경을 부모 색상으로 채움
            if (Parent != null)
            {
                using (SolidBrush parentBrush = new SolidBrush(Parent.BackColor))
                    pevent.Graphics.FillRectangle(parentBrush, ClientRectangle);
            }
            else
            {
                pevent.Graphics.Clear(Color.Transparent);
            }

            //int pressOffset = isPressed ? 2 : 0;
            //int currentShadowExpand = isHover ? hoverShadowExpand : 0;
            //int currentShadowBlur = shadowBlur + currentShadowExpand;
            int pressOffset = 0;              // 버튼 위치 고정
            int currentShadowExpand = 0;      // 그림자 크기 고정
            int currentShadowBlur = shadowBlur; // 기본 blur만 사용

            // 그림자
            if (shadowOffset > 0)
                DrawSoftShadow(pevent.Graphics, pressOffset, currentShadowExpand, currentShadowBlur);

            // 버튼 영역
            int buttonOffset = 2;
            Rectangle rect = new Rectangle(
                depthOffset,
                depthOffset + buttonOffset + pressOffset,
                Width - shadowOffset - depthOffset - currentShadowExpand,
                Height - shadowOffset - depthOffset - currentShadowExpand - buttonOffset
            );

            using (GraphicsPath path = GetRoundedRectangle(rect, cornerRadius))
            {
                // 배경
                using (SolidBrush brush = new SolidBrush(buttonBackColor))
                    pevent.Graphics.FillPath(brush, path);

                // border
                using (Pen pen = new Pen(isHover ? hoverBorderColor : borderColor, borderSize))
                    pevent.Graphics.DrawPath(pen, path);
            }

            // 텍스트
            TextRenderer.DrawText(pevent.Graphics, Text, Font, rect, textColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
        }

        private void DrawSoftShadow(Graphics graphics, int pressOffset, int currentShadowExpand, int currentShadowBlur)
        {
            int layers = Math.Max(3, currentShadowBlur / 2);

            for (int i = 0; i < layers; i++)
            {
                float layerOpacity = (1.0f - (float)i / layers) * 0.3f;
                int layerOffset = shadowOffset + (i * currentShadowBlur / layers);
                int layerExpand = i * (currentShadowExpand + 1) / layers;

                Rectangle shadowRect = new Rectangle(
                    depthOffset + layerOffset - layerExpand,
                    depthOffset + layerOffset + pressOffset - layerExpand,
                    Width - depthOffset - 1 + (layerExpand * 2),
                    Height - depthOffset - 1 + (layerExpand * 2)
                );

                using (GraphicsPath shadowPath = GetRoundedRectangle(shadowRect, cornerRadius + layerExpand + 2))
                using (PathGradientBrush shadowBrush = new PathGradientBrush(shadowPath))
                {
                    Color centerColor = Color.FromArgb(
                        (int)(shadowColor.A * layerOpacity),
                        shadowColor.R, shadowColor.G, shadowColor.B
                    );

                    Color edgeColor = Color.FromArgb(0, shadowColor.R, shadowColor.G, shadowColor.B);

                    shadowBrush.CenterColor = centerColor;
                    shadowBrush.SurroundColors = new[] { edgeColor };
                    shadowBrush.CenterPoint = new PointF(
                        shadowRect.Left + shadowRect.Width / 2f,
                        shadowRect.Top + shadowRect.Height / 2.5f
                    );

                    Blend blend = new Blend
                    {
                        Factors = new float[] { 0f, 0.3f, 0.7f, 1f },
                        Positions = new float[] { 0f, 0.3f, 0.7f, 1f }
                    };
                    shadowBrush.Blend = blend;

                    graphics.FillPath(shadowBrush, shadowPath);
                }
            }
        }

        private GraphicsPath GetRoundedRectangle(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            if (radius <= 0)
            {
                path.AddRectangle(rect);
                return path;
            }

            int diameter = radius * 2;
            Rectangle arc = new Rectangle(rect.Location, new Size(diameter, diameter));

            // 왼쪽 위
            path.AddArc(arc, 180, 90);
            // 오른쪽 위
            arc.X = rect.Right - diameter;
            path.AddArc(arc, 270, 90);
            // 오른쪽 아래
            arc.Y = rect.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            // 왼쪽 아래
            arc.X = rect.Left;
            path.AddArc(arc, 90, 90);

            path.CloseAllFigures();
            return path;
        }
    }
}

using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace count_dead_sign
{
    public class CircularProgressBar : UserControl
    {
        private int _value = 0;
        private int _maxValue = 100;
        private Color _progressColor = Color.DodgerBlue;
        private Color _backColor = Color.LightGray;
        private int _lineWidth = 10;

        [Category("Behavior")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public int Value
        {
            get => _value;
            set
            {
                if (value < 0) value = 0;
                if (value > _maxValue) value = _maxValue;
                _value = value;
                Invalidate();
            }
        }

        [Category("Behavior")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public int Maximum
        {
            get => _maxValue;
            set
            {
                if (value <= 0) value = 1;
                _maxValue = value;
                Invalidate();
            }
        }

        [Category("Appearance")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public Color ProgressColor
        {
            get => _progressColor;
            set { _progressColor = value; Invalidate(); }
        }

        [Category("Appearance")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public Color CircleBackColor
        {
            get => _backColor;
            set { _backColor = value; Invalidate(); }
        }

        [Category("Appearance")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        public int LineWidth
        {
            get => _lineWidth;
            set { _lineWidth = Math.Max(1, value); Invalidate(); }
        }

        public CircularProgressBar()
        {
            this.DoubleBuffered = true;
            this.Size = new Size(100, 100);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // 원형 영역
            Rectangle rect = new Rectangle(
                _lineWidth,
                _lineWidth,
                this.Width - _lineWidth * 2,
                this.Height - _lineWidth * 2
            );

            // 배경 원
            using (Pen backPen = new Pen(_backColor, _lineWidth))
            {
                e.Graphics.DrawArc(backPen, rect, -90, 360);
            }

            // 진행 원
            float sweepAngle = (float)_value / _maxValue * 360f;
            using (Pen progressPen = new Pen(_progressColor, _lineWidth))
            {
                e.Graphics.DrawArc(progressPen, rect, -90, sweepAngle);
            }

            // 가운데 % 텍스트
            string percent = $"{(int)((float)_value / _maxValue * 100)}%";
            using (Font font = new Font(this.Font.FontFamily, 14, FontStyle.Bold))
            using (Brush brush = new SolidBrush(Color.White))
            {
                SizeF textSize = e.Graphics.MeasureString(percent, font);
                e.Graphics.DrawString(
                    percent,
                    font,
                    brush,
                    (this.Width - textSize.Width) / 2,
                    (this.Height - textSize.Height) / 2
                );
            }
        }
    }
}

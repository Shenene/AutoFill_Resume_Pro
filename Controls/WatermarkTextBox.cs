using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace AutoFill_Resume_Pro.Controls
{
    public class WatermarkTextBox : TextBox
    {
        private string watermarkText = string.Empty;

        [Browsable(true)]
        [Category("Appearance")]
        [Description("Watermark text to display when the textbox is empty.")]
        public string WatermarkText
        {
            get { return watermarkText; }
            set { watermarkText = value; Invalidate(); }
        }

        public WatermarkTextBox()
        {
            // Ensure we repaint the control when necessary.
            this.SetStyle(ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            // Draw the normal text.
            TextRenderer.DrawText(e.Graphics, this.Text, this.Font, this.ClientRectangle, this.ForeColor);

            // If no text and not focused, draw the watermark.
            if (string.IsNullOrEmpty(this.Text) && !this.Focused && !string.IsNullOrEmpty(WatermarkText))
            {
                TextFormatFlags flags = TextFormatFlags.NoPadding | TextFormatFlags.Top | TextFormatFlags.Left;
                Rectangle rect = new Rectangle(1, 1, this.Width - 2, this.Height - 2);
                TextRenderer.DrawText(e.Graphics, WatermarkText, this.Font, rect, Color.Gray, flags);
            }
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            Invalidate(); // Redraw to remove watermark if text is entered.
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            Invalidate(); // Remove watermark when focused.
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            Invalidate(); // Redraw watermark when focus is lost.
        }
    }
}

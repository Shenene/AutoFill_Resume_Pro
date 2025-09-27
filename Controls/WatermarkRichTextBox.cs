using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace AutoFill_Resume_Pro.Controls
{
    public class WatermarkRichTextBox : RichTextBox
    {
        private string watermarkText = string.Empty;
        private bool isSingleLine = false;

        [Browsable(true)]
        [Category("Appearance")]
        [Description("Watermark text to display when the control is empty.")]
        public string WatermarkText
        {
            get { return watermarkText; }
            set { watermarkText = value; Invalidate(); }
        }

        [Browsable(true)]
        [Category("Behavior")]
        [Description("Set to true for single-line mode.")]
        public bool IsSingleLine
        {
            get { return isSingleLine; }
            set
            {
                isSingleLine = value;
                if (isSingleLine)
                {
                    // Simulate single-line mode:
                    this.Multiline = false;
                    this.ScrollBars = RichTextBoxScrollBars.None;
                    // Optionally adjust the height to roughly fit one line.
                    this.Height = this.Font.Height + 6;
                }
                else
                {
                    this.Multiline = true;
                    this.ScrollBars = RichTextBoxScrollBars.Vertical;
                }
                Invalidate();
            }
        }

        public WatermarkRichTextBox()
        {
            // Set default font to Calibri, 12pt.
            this.Font = new Font("Calibri", 12F);
            // Set default user input colors.
            this.ForeColor = Color.FromArgb(0, 0, 0);       // Black text for user input.
            this.BackColor = Color.FromArgb(255, 255, 255);   // White background for user input.
            // Enable double buffering for smoother painting.
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
        }

        protected override void OnTextChanged(EventArgs e)
        {
            base.OnTextChanged(e);
            Invalidate(); // Redraw immediately when text changes.
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            Invalidate(); // Redraw immediately on key up.
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            Invalidate(); // Redraw on mouse up in case text is cleared.
        }

        protected override void OnContentsResized(ContentsResizedEventArgs e)
        {
            base.OnContentsResized(e);
            Invalidate(); // Redraw if content size changes.
        }

        protected override void OnGotFocus(EventArgs e)
        {
            base.OnGotFocus(e);
            // When focused, use user input colors.
            this.ForeColor = Color.FromArgb(0, 0, 0);       // Black text.
            this.BackColor = Color.FromArgb(255, 255, 255);   // White background.
            Invalidate();
        }

        protected override void OnLostFocus(EventArgs e)
        {
            base.OnLostFocus(e);
            // When not focused and empty, use watermark background.
            if (string.IsNullOrEmpty(this.Text))
            {
                this.BackColor = Color.FromArgb(240, 240, 240); // Light gray background.
            }
            Invalidate();
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_PAINT = 0x000F;
            base.WndProc(ref m);

            // After default painting, if there's no text and a watermark is defined, draw the watermark.
            if (m.Msg == WM_PAINT)
            {
                if (string.IsNullOrEmpty(this.Text) && !string.IsNullOrEmpty(WatermarkText) && !this.Focused)
                {
                    using (Graphics g = this.CreateGraphics())
                    {
                        using (StringFormat sf = new StringFormat())
                        {
                            sf.Alignment = StringAlignment.Near;
                            // For single-line, center vertically; for multiline, align near the top.
                            sf.LineAlignment = isSingleLine ? StringAlignment.Center : StringAlignment.Near;
                            sf.FormatFlags = StringFormatFlags.LineLimit;
                            using (SolidBrush brush = new SolidBrush(Color.FromArgb(150, 150, 150))) // Light gray watermark text.
                            {
                                g.DrawString(WatermarkText, this.Font, brush, this.ClientRectangle, sf);
                            }
                        }
                    }
                }
            }
        }
    }
}

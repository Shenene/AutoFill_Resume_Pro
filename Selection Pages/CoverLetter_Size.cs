using AutoFill_Resume_Pro.Forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoFill_Resume_Pro.Selection_Pages
{
    public partial class CoverLetter_Size: Form
    {
        public CoverLetter_Size()
        {
            this.AutoScaleMode = AutoScaleMode.None;

            // Set consistent form appearance
            this.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Regular, GraphicsUnit.Point);

            InitializeComponent();

            this.ClientSize = new Size(800, 950);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(800, 768);
            this.MaximizeBox = false;
            this.DoubleBuffered = true;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 874);

            this.Paint += Resume_Size_Paint;
        }

        public CoverLetter_Size(Size parentSize)
        {
            InitializeComponent();
            this.Size = parentSize;
        }

        private void Resume_Size_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            var f1 = new Form1
            {
                Size = this.Size
            };
            f1.Show();

            // Close this form (we won't come back to it)
            this.Close();
        }

        private void button_A4Format_Click(object sender, EventArgs e)
        {
            var a4 = new A4_CoverLetter_Info
            {
                Size = this.Size
            };
            a4.Show();
            this.Close();
        }

        private void buttonUSLetterFormat_Click(object sender, EventArgs e)
        {
            var us = new US_CoverLetter_Info
            {
                Size = this.Size
            };
            us.Show();
            this.Close();
        }
    }
}
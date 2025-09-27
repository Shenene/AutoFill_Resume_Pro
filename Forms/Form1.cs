using AutoFill_Resume_Pro.Selection_Pages;
using AutoFill_Resume_Pro.Templates;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AutoFill_Resume_Pro
{
    public partial class Form1 : Form
    {
        private Panel panel_Main; // Renamed to avoid ambiguity
        private Label label_TRE_Main; // Renamed to avoid ambiguity
        private string[] userData; // Store user data here

        // Field to store previous size if passed in.
        private Size _previousSize = Size.Empty;

        // New fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        // Add this helper method to collect the latest user data.
        public string[] GetUserData()
        {
            return new string[]
            {
                textBox_FN.Text,  // Full Name
                textBox_PT.Text,  // Position Title
                textBox_P.Text,   // Phone
                textBox_EM.Text,  // Email
                textBox_L.Text,   // Location
                textBox_LP.Text   // LinkedIn URL
        };
    }

        // Default constructor: fresh launch, force size to 800x950.
        public Form1()
        {
            InitializeComponent();
            InitializeFormSettings(new Size(800, 950));  // Use default settings
            LoadUserSettings();  // Load settings
            this.Size = new Size(800, 950); // Force default overall window size.

            // Defer watermark setting until after the form is shown.
            this.Shown += Form1_Shown;
        }

        // Parameterized constructor: when returning, pass the current size.
        public Form1(Size previousSize)
        {
            InitializeComponent();
            _previousSize = previousSize; // Store passed size.
            InitializeFormSettings(previousSize);
            LoadUserSettings();
            this.Size = _previousSize; // Force the form to the passed size.
            this.Shown += Form1_Shown;
        }

        private void InitializeFormSettings(Size formSize)
        {
            this.AutoScaleMode = AutoScaleMode.None;  // Disable auto-scaling
            this.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Regular, GraphicsUnit.Point);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 874);
            this.DoubleBuffered = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Background Panel Setup
            panel_Background.Dock = DockStyle.Fill;

            // Panel_Main Setup (inside panel_Background)
            panel_Main = new Panel
            {
                Dock = DockStyle.Fill
            };
            panel_Background.Controls.Add(panel_Main);

            label_TRE_Main = new Label
            {
                Text = "The Resume Edge",
                Font = new Font("Arial", 12, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(10, 10)
            };
            panel_Main.Controls.Add(label_TRE_Main);

            // PictureBox_PC Setup (for your paperclip image)
            pictureBox_PC.BringToFront();
            pictureBox_PC.Size = new Size(this.Width / 9, this.Width / 9);
            pictureBox_PC.Image = ResizeImage(Properties.Resources.PaperClips, pictureBox_PC.Width, pictureBox_PC.Height);
            pictureBox_PC.SizeMode = PictureBoxSizeMode.Zoom;

            // Logo Setup (pictureBox_Logo in panel_TC)
            pictureBox_Logo.BringToFront();
            pictureBox_Logo.Size = new Size(200, 100);
            pictureBox_Logo.Image = Properties.Resources.TRE_Logo;
            pictureBox_Logo.SizeMode = PictureBoxSizeMode.Zoom;

            // Center the logo on form load.
            CenterLogo();

            // Subscribe to TextChanged events for autosave detection.
            textBox_FN.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_P.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_L.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_LP.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes (adjust as needed)
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Make sure FormClosing event is subscribed.
            InitializeFormClosingEvent();

            // Set the ActiveControl to avoid radioButton1 getting focus.
            this.ActiveControl = textBox_FN;
        }

        /// <summary>
        /// Event handler for the Shown event to set watermarks after the form is fully created.
        /// </summary>
        private void Form1_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // Now set the watermarks. Since the handles are created, these should work.
                WatermarkHelper.SetCue(textBox_FN, " AIDEN WALSH");
                WatermarkHelper.SetCue(textBox_PT, " POSITION TITLE");
                WatermarkHelper.SetCue(textBox_P, " (212) 555-1234");
                WatermarkHelper.SetCue(textBox_EM, " youremail@mail.com");
                WatermarkHelper.SetCue(textBox_L, " New York, NY");
                WatermarkHelper.SetCue(textBox_LP, " linkedin.com/in/username");
            }
        }

        /// <summary>
        /// Helper method to determine design mode reliably.
        /// </summary>
        private bool IsInDesignMode()
        {
            return (LicenseManager.UsageMode == LicenseUsageMode.Designtime) || this.DesignMode;
        }

        private void AutosaveTimer_Tick(object sender, EventArgs e)
        {
            if (unsavedChanges)
            {
                AutosaveUserSettings();
                // Optionally: update a status label to indicate an autosave.
            }
        }

        // Helper method to perform autosave silently.
        private void AutosaveUserSettings()
        {
            Properties.Settings.Default.Form1_FullName = textBox_FN.Text;
            Properties.Settings.Default.Form1_ProTitle = textBox_PT.Text;
            Properties.Settings.Default.Form1_Phone = textBox_P.Text;
            Properties.Settings.Default.Form1_Email = textBox_EM.Text;
            Properties.Settings.Default.Form1_Location = textBox_L.Text;
            Properties.Settings.Default.Form1_LinkedIn = textBox_LP.Text;
            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        private void StoreUserData(string[] data)
        {
            userData = data;
        }

        // Call this method when opening the resume template.
        private void OpenResumeTemplate()
        {
            if (userData != null)
            {
                A4_Resume resumeForm = new A4_Resume(userData);
                resumeForm.Show();
            }
            else
            {
                MessageBox.Show("Please complete all steps before opening the resume.");
            }
        }

        // Center the logo without interrupting the layout.
        private void CenterLogo()
        {
            if (pictureBox_Logo != null && panel_TC != null)
            {
                this.SuspendLayout();
                pictureBox_Logo.Location = new Point(
                    (panel_TC.ClientSize.Width - pictureBox_Logo.Width) / 2,
                    (panel_TC.ClientSize.Height - pictureBox_Logo.Height) / 2);
                this.ResumeLayout();
            }
        }

        // Override OnResize to handle resizing smoothly.
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            this.SuspendLayout();
            this.Invalidate();
            this.ResumeLayout();
        }

        // Method to resize an image without losing quality.
        private Image ResizeImage(Image img, int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.DrawImage(img, 0, 0, width, height);
            }
            return bmp;
        }

        // High-quality rendering for the entire form.
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            base.OnPaint(e);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // If there are unsaved changes, autosave without prompting.
            if (unsavedChanges)
            {
                AutosaveUserSettings();
            }
            // Optional cleanup here.
        }

        private void InitializeFormClosingEvent()
        {
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
        }

        // Example button click events for navigation between forms.
        private void button1_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var RS = new Resume_Size();
            RS.Size = this.Size;
            RS.Show();

            // Hide Form1 so the app stays alive
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var CLS = new CoverLetter_Size();
            CLS.Size = this.Size;
            CLS.Show();

            // Hide Form1 so the app stays alive
            this.Hide();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var RLS = new ReferenceLetter_Size();
            RLS.Size = this.Size;
            RLS.Show();

            // Hide Form1 so the app stays alive
            this.Hide();
        }

        // Save button event handler.
        private void button_Save_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Form1_FullName = textBox_FN.Text;
            Properties.Settings.Default.Form1_ProTitle = textBox_PT.Text;
            Properties.Settings.Default.Form1_Phone = textBox_P.Text;
            Properties.Settings.Default.Form1_Email = textBox_EM.Text;
            Properties.Settings.Default.Form1_Location = textBox_L.Text;
            Properties.Settings.Default.Form1_LinkedIn = textBox_LP.Text;
            Properties.Settings.Default.Save();
            unsavedChanges = false;
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Load button event handler.
        private void button_Load_Click(object sender, EventArgs e)
        {
            LoadUserSettings();
        }

        // Helper method to load user settings into the TextBoxes.
        private void LoadUserSettings()
        {
            textBox_FN.Text = Properties.Settings.Default.Form1_FullName;
            textBox_PT.Text = Properties.Settings.Default.Form1_ProTitle;
            textBox_P.Text = Properties.Settings.Default.Form1_Phone;
            textBox_EM.Text = Properties.Settings.Default.Form1_Email;
            textBox_L.Text = Properties.Settings.Default.Form1_Location;
            textBox_LP.Text = Properties.Settings.Default.Form1_LinkedIn;
        }
    }

    // WatermarkHelper is now defined after Form1 so that the designer loads Form1 as the first class.
    public static class WatermarkHelper
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wp, string lp);

        private const int EM_SETCUEBANNER = 0x1501;

        public static void SetCue(TextBox textBox, string cueText)
        {
            if (textBox.IsHandleCreated)
            {
                SendMessage(textBox.Handle, EM_SETCUEBANNER, (IntPtr)1, cueText);
            }
        }
    }
}

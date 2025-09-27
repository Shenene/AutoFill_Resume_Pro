using AutoFill_Resume_Pro.Templates;
using System;
using System.ComponentModel;         // For LicenseUsageMode
using System.Drawing;
using System.Windows.Forms;
using AutoFill_Resume_Pro.Selection_Pages;
using System.IO;                         // For Path.Combine
using System.Text.RegularExpressions;    // For Regex
using System.Collections.Generic;        // For List<T>

namespace AutoFill_Resume_Pro.Forms
{
    public partial class A4_Reference_Info : Form
    {
        // Fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        // helper method to collect the latest user data
        public string[] GetUserData()
        {
            return new string[]
            {
                textBox_RN1.Text,   // Reference Name 1
                textBox_PT1.Text,   // Position Title 1
                textBox_ORG1.Text,  // Organization 1
                textBox_PH1.Text,   // Phone 1
                textBox_EM1.Text,   // Email 1
                textBox_RN2.Text,   // Reference Name 2
                textBox_PT2.Text,   // Position Title 2
                textBox_ORG2.Text,  // Organization 2
                textBox_PH2.Text,   // Phone 2
                textBox_EM2.Text,   // Email 2
                textBox_RN3.Text,   // Reference Name 3
                textBox_PT3.Text,   // Position Title 3
                textBox_ORG3.Text,  // Organization 3
                textBox_PH3.Text,   // Phone 3
                textBox_EM3.Text,   // Email 3
                textBox_RN4.Text,   // Reference Name 4
                textBox_PT4.Text,   // Position Title 4
                textBox_ORG4.Text,  // Organization 4
                textBox_PH4.Text,   // Phone 4
                textBox_EM4.Text,   // Email 4
                textBox_RN5.Text,   // Reference Name 5
                textBox_PT5.Text,   // Position Title 5
                textBox_ORG5.Text,  // Organization 5
                textBox_PH5.Text,   // Phone 5
                textBox_EM5.Text    // Email 5                
            };
        }

        public A4_Reference_Info()
        {
            this.AutoScaleMode = AutoScaleMode.None;
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

            // Subscribe to form events.
            this.Load += A4_Reference_Info_Load;
            this.Shown += A4_Reference_Info_Shown;
            this.FormClosing += A4_Reference_Info_FormClosing;
        }

        public A4_Reference_Info(Size parentSize)
        {
            InitializeComponent();
            this.Size = parentSize;
            this.Load += A4_Reference_Info_Load;
            this.Shown += A4_Reference_Info_Shown;
            this.FormClosing += A4_Reference_Info_FormClosing;
        }

        private void A4_Reference_Info_Load(object sender, EventArgs e)
        {
            // Subscribe to TextChanged events for each reference textbox.
            // Reference 1
            textBox_RN1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_ORG1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PH1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM1.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Reference 2
            textBox_RN2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_ORG2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PH2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM2.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Reference 3
            textBox_RN3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_ORG3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PH3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM3.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Reference 4
            textBox_RN4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_ORG4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PH4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM4.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Reference 5
            textBox_RN5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PT5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_ORG5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_PH5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_EM5.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes; adjust as needed.
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Load previously saved reference data.
            LoadUserSettings();
        }

        private void A4_Reference_Info_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // Set watermark placeholders directly using custom control properties.
                // Reference 1
                textBox_RN1.IsSingleLine = true;
                textBox_RN1.WatermarkText = "Enter Reference Name";
                textBox_PT1.IsSingleLine = true;
                textBox_PT1.WatermarkText = "Enter Position Title";
                textBox_ORG1.IsSingleLine = true;
                textBox_ORG1.WatermarkText = "Enter Organization";
                textBox_PH1.IsSingleLine = true;
                textBox_PH1.WatermarkText = "Enter Phone No";
                textBox_EM1.IsSingleLine = true;
                textBox_EM1.WatermarkText = "Enter Email";
                // Reference 2
                textBox_RN2.IsSingleLine = true;
                textBox_RN2.WatermarkText = "Enter Reference Name";
                textBox_PT2.IsSingleLine = true;
                textBox_PT2.WatermarkText = "Enter Position Title";
                textBox_ORG2.IsSingleLine = true;
                textBox_ORG2.WatermarkText = "Enter Organization";
                textBox_PH2.IsSingleLine = true;
                textBox_PH2.WatermarkText = "Enter Phone No";
                textBox_EM2.IsSingleLine = true;
                textBox_EM2.WatermarkText = "Enter Email";
                // Reference 3
                textBox_RN3.IsSingleLine = true;
                textBox_RN3.WatermarkText = "Enter Reference Name";
                textBox_PT3.IsSingleLine = true;
                textBox_PT3.WatermarkText = "Enter Position Title";
                textBox_ORG3.IsSingleLine = true;
                textBox_ORG3.WatermarkText = "Enter Organization";
                textBox_PH3.IsSingleLine = true;
                textBox_PH3.WatermarkText = "Enter Phone No";
                textBox_EM3.IsSingleLine = true;
                textBox_EM3.WatermarkText = "Enter Email";
                // Reference 4
                textBox_RN4.IsSingleLine = true;
                textBox_RN4.WatermarkText = "Enter Reference Name";
                textBox_PT4.IsSingleLine = true;
                textBox_PT4.WatermarkText = "Enter Position Title";
                textBox_ORG4.IsSingleLine = true;
                textBox_ORG4.WatermarkText = "Enter Organization";
                textBox_PH4.IsSingleLine = true;
                textBox_PH4.WatermarkText = "Enter Phone No";
                textBox_EM4.IsSingleLine = true;
                textBox_EM4.WatermarkText = "Enter Email";
                // Reference 5
                textBox_RN5.IsSingleLine = true;
                textBox_RN5.WatermarkText = "Enter Reference Name";
                textBox_PT5.IsSingleLine = true;
                textBox_PT5.WatermarkText = "Enter Position Title";
                textBox_ORG5.IsSingleLine = true;
                textBox_ORG5.WatermarkText = "Enter Organization";
                textBox_PH5.IsSingleLine = true;
                textBox_PH5.WatermarkText = "Enter Phone No";
                textBox_EM5.IsSingleLine = true;
                textBox_EM5.WatermarkText = "Enter Email";
            }
        }

        private bool IsInDesignMode()
        {
            return (LicenseManager.UsageMode == LicenseUsageMode.Designtime) || this.DesignMode;
        }

        private void AutosaveTimer_Tick(object sender, EventArgs e)
        {
            if (unsavedChanges)
            {
                AutosaveUserSettings();
                // Optionally, update a status label here.
            }
        }

        // Saves cover letter data into user settings.
        private void AutosaveUserSettings()
        {
            // Reference 1
            Properties.Settings.Default.A4_Ref_RefName1 = textBox_RN1.Text;
            Properties.Settings.Default.A4_Ref_PositionT1 = textBox_PT1.Text;
            Properties.Settings.Default.A4_Ref_Org1 = textBox_ORG1.Text;
            Properties.Settings.Default.A4_Ref_Phone1 = textBox_PH1.Text;
            Properties.Settings.Default.A4_Ref_Email1 = textBox_EM1.Text;
            // Reference 2
            Properties.Settings.Default.A4_Ref_RefName2 = textBox_RN2.Text;
            Properties.Settings.Default.A4_Ref_PositionT2 = textBox_PT2.Text;
            Properties.Settings.Default.A4_Ref_Org2 = textBox_ORG2.Text;
            Properties.Settings.Default.A4_Ref_Phone2 = textBox_PH2.Text;
            Properties.Settings.Default.A4_Ref_Email2 = textBox_EM2.Text;
            // Reference 3
            Properties.Settings.Default.A4_Ref_RefName3 = textBox_RN3.Text;
            Properties.Settings.Default.A4_Ref_PositionT3 = textBox_PT3.Text;
            Properties.Settings.Default.A4_Ref_Org3 = textBox_ORG3.Text;
            Properties.Settings.Default.A4_Ref_Phone3 = textBox_PH3.Text;
            Properties.Settings.Default.A4_Ref_Email3 = textBox_EM3.Text;
            // Reference 4
            Properties.Settings.Default.A4_Ref_RefName4 = textBox_RN4.Text;
            Properties.Settings.Default.A4_Ref_PositionT4 = textBox_PT4.Text;
            Properties.Settings.Default.A4_Ref_Org4 = textBox_ORG4.Text;
            Properties.Settings.Default.A4_Ref_Phone4 = textBox_PH4.Text;
            Properties.Settings.Default.A4_Ref_Email4 = textBox_EM4.Text;
            // Reference 5
            Properties.Settings.Default.A4_Ref_RefName5 = textBox_RN5.Text;
            Properties.Settings.Default.A4_Ref_PositionT5 = textBox_PT5.Text;
            Properties.Settings.Default.A4_Ref_Org5 = textBox_ORG5.Text;
            Properties.Settings.Default.A4_Ref_Phone5 = textBox_PH5.Text;
            Properties.Settings.Default.A4_Ref_Email5 = textBox_EM5.Text;

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        // Loads saved cover letter data into controls.
        private void LoadUserSettings()
        {
            // Reference 1
            textBox_RN1.Text = Properties.Settings.Default.A4_Ref_RefName1;
            textBox_PT1.Text = Properties.Settings.Default.A4_Ref_PositionT1;
            textBox_ORG1.Text = Properties.Settings.Default.A4_Ref_Org1;
            textBox_PH1.Text = Properties.Settings.Default.A4_Ref_Phone1;
            textBox_EM1.Text = Properties.Settings.Default.A4_Ref_Email1;
            // Reference 2
            textBox_RN2.Text = Properties.Settings.Default.A4_Ref_RefName2;
            textBox_PT2.Text = Properties.Settings.Default.A4_Ref_PositionT2;
            textBox_ORG2.Text = Properties.Settings.Default.A4_Ref_Org2;
            textBox_PH2.Text = Properties.Settings.Default.A4_Ref_Phone2;
            textBox_EM2.Text = Properties.Settings.Default.A4_Ref_Email2;
            // Reference 3
            textBox_RN3.Text = Properties.Settings.Default.A4_Ref_RefName3;
            textBox_PT3.Text = Properties.Settings.Default.A4_Ref_PositionT3;
            textBox_ORG3.Text = Properties.Settings.Default.A4_Ref_Org3;
            textBox_PH3.Text = Properties.Settings.Default.A4_Ref_Phone3;
            textBox_EM3.Text = Properties.Settings.Default.A4_Ref_Email3;
            // Reference 4
            textBox_RN4.Text = Properties.Settings.Default.A4_Ref_RefName4;
            textBox_PT4.Text = Properties.Settings.Default.A4_Ref_PositionT4;
            textBox_ORG4.Text = Properties.Settings.Default.A4_Ref_Org4;
            textBox_PH4.Text = Properties.Settings.Default.A4_Ref_Phone4;
            textBox_EM4.Text = Properties.Settings.Default.A4_Ref_Email4;
            // Reference 5
            textBox_RN5.Text = Properties.Settings.Default.A4_Ref_RefName5;
            textBox_PT5.Text = Properties.Settings.Default.A4_Ref_PositionT5;
            textBox_ORG5.Text = Properties.Settings.Default.A4_Ref_Org5;
            textBox_PH5.Text = Properties.Settings.Default.A4_Ref_Phone5;
            textBox_EM5.Text = Properties.Settings.Default.A4_Ref_Email5;
        }

        private void A4_Reference_Info_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();
        }

        private void Resume_Size_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        }

        // Navigation: Return to Form1.
        private void button_Return_Click_1(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var f1 = new Form1
            {
                Size = this.Size
            };
            f1.Show();

            // Close this form (we won't come back to it)
            this.Close();
        }

        // Navigation: Go to the References Template.
        private void button_GoTo_Click(object sender, EventArgs e)
        {
            // 1) First flush any unsaved edits into Settings:
            if (unsavedChanges)
                AutosaveUserSettings();
            Properties.Settings.Default.Save();

            // Retrieve the data saved from Form1
            string[] userData = new string[]
                {
            // Data from Form1 (inicates 0-5)
            Properties.Settings.Default.Form1_FullName,
                Properties.Settings.Default.Form1_ProTitle,
                Properties.Settings.Default.Form1_Phone,
                Properties.Settings.Default.Form1_Email,
                Properties.Settings.Default.Form1_Location,
                Properties.Settings.Default.Form1_LinkedIn,         
                // Data from CoverLetter (indicates 6 - 30)
                Properties.Settings.Default.A4_Ref_RefName1,
                Properties.Settings.Default.A4_Ref_PositionT1,
                Properties.Settings.Default.A4_Ref_Org1,
                Properties.Settings.Default.A4_Ref_Phone1,
                Properties.Settings.Default.A4_Ref_Email1,
                Properties.Settings.Default.A4_Ref_RefName2,
                Properties.Settings.Default.A4_Ref_PositionT2,
                Properties.Settings.Default.A4_Ref_Org2,
                Properties.Settings.Default.A4_Ref_Phone2,
                Properties.Settings.Default.A4_Ref_Email2,
                Properties.Settings.Default.A4_Ref_RefName3,
                Properties.Settings.Default.A4_Ref_PositionT3,
                Properties.Settings.Default.A4_Ref_Org3,
                Properties.Settings.Default.A4_Ref_Phone3,
                Properties.Settings.Default.A4_Ref_Email3,
                Properties.Settings.Default.A4_Ref_RefName4,
                Properties.Settings.Default.A4_Ref_PositionT4,
                Properties.Settings.Default.A4_Ref_Org4,
                Properties.Settings.Default.A4_Ref_Phone4,
                Properties.Settings.Default.A4_Ref_Email4,
                Properties.Settings.Default.A4_Ref_RefName5,
                Properties.Settings.Default.A4_Ref_PositionT5,
                Properties.Settings.Default.A4_Ref_Org5,
                Properties.Settings.Default.A4_Ref_Phone5,
                Properties.Settings.Default.A4_Ref_Email5,
        };

            var a4_References = new A4_References(userData)
            {
                Size = this.Size
            };
            a4_References.Show();

            // Close the “info” form
            this.Close();
        }

        // Navigation: Go back to ReferenceLetter_Size.
        private void button_Back_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var rls = new ReferenceLetter_Size
            {
                Size = this.Size
            };
            rls.Show();

            // Close this form (we're not returning to it)
            this.Close();
        }

        // Manual Save button event handler.
        private void button_Save_A4Ref_Click_1(object sender, EventArgs e)
        {
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

using AutoFill_Resume_Pro.Templates;
// using AutoFill_Resume_Pro.Helpers;  // Not needed anymore with the custom control
using AutoFill_Resume_Pro.Selection_Pages;
using System;
using System.ComponentModel;         // For LicenseUsageMode
using System.Drawing;
using System.Windows.Forms;
using System.IO;                         // For Path.Combine
using System.Text.RegularExpressions;    // For Regex
using System.Collections.Generic;        // For List<T>
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using System.Windows.Forms.VisualStyles;

namespace AutoFill_Resume_Pro.Forms
{
    public partial class US_CoverLetter_Info : Form
    {
        // Fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        // helper method to collect the latest user data
        public string[] GetUserData()
        {
            return new string[]
            {
                textBox_D.Text,     // Date
                textBox_HMFN.Text,  // Hiring Manager Full Name
                textBox_CN.Text,    // Company Name
                textBox_CA.Text,    // Company Address
                textBox_City.Text,  // City
                textBox_HM.Text,    // Hiring Manager
                textBox_B.Text,     // Body
                textBox_FN.Text     // Your Full Name
            };
        }

        public US_CoverLetter_Info()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            this.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Regular, GraphicsUnit.Point);
            InitializeComponent();

            // Set the WatermarkText for textBox_B (multiline control).
            textBox_B.WatermarkText = @"Introduction: Introduce yourself with your name and a brief professional background. State the position you're applying for, where you found the job listing, and why you're interested in this position and the company.
Highlight your ability to apply your knowledge in practical settings, collaborate effectively with diverse teams, and take initiative.
Highlight Qualifications and Experience: Summarize your relevant qualifications, including education and certifications. Highlight your professional experience with roles and achievements that align with the job description, providing specific examples.

Show Enthusiasm and Fit: Express your enthusiasm for the job and the company. Mention specific aspects of the company that appeal to you and explain how the opportunity fits with your career goals. Highlight your strengths and how they align with the company's needs.

Conclusion and Call to Action: Thank the hiring manager for their time and consideration. Invite them to review your resume for more details and encourage them to contact you to schedule an interview. Mention your eagerness to discuss how you can contribute to the company's success.
Editing and Proofreading: Thoroughly edit and proofread your cover letter to ensure there are no errors. Ensure the letter is clear, concise, and tailored to the specific job and company.

Yours sincerely";

            this.ClientSize = new Size(800, 950);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(800, 768);
            this.MaximizeBox = false;
            this.DoubleBuffered = true;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 874);

            this.Paint += Resume_Size_Paint;

            // Subscribe to form events.
            this.Load += US_CoverLetter_Info_Load;
            this.Shown += US_CoverLetter_Info_Shown;
            this.FormClosing += US_CoverLetter_Info_FormClosing;
        }

        public US_CoverLetter_Info(Size parentSize)
        {
            InitializeComponent();
            // Optionally, set the watermark for textBox_B here if not set in the designer.
            textBox_B.WatermarkText = @"Introduction: Introduce yourself with your name and a brief professional background. State the position you're applying for, where you found the job listing, and why you're interested in this position and the company.
Highlight your ability to apply your knowledge in practical settings, collaborate effectively with diverse teams, and take initiative.
Highlight Qualifications and Experience: Summarize your relevant qualifications, including education and certifications. Highlight your professional experience with roles and achievements that align with the job description, providing specific examples.

Show Enthusiasm and Fit: Express your enthusiasm for the job and the company. Mention specific aspects of the company that appeal to you and explain how the opportunity fits with your career goals. Highlight your strengths and how they align with the company's needs.

Conclusion and Call to Action: Thank the hiring manager for their time and consideration. Invite them to review your resume for more details and encourage them to contact you to schedule an interview. Mention your eagerness to discuss how you can contribute to the company's success.
Editing and Proofreading: Thoroughly edit and proofread your cover letter to ensure there are no errors. Ensure the letter is clear, concise, and tailored to the specific job and company.

Yours sincerely";

            this.Size = parentSize;
            this.Load += US_CoverLetter_Info_Load;
            this.Shown += US_CoverLetter_Info_Shown;
            this.FormClosing += US_CoverLetter_Info_FormClosing;
        }

        private void US_CoverLetter_Info_Load(object sender, EventArgs e)
        {
            // Subscribe to TextChanged events for autosave detection.
            textBox_D.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_HMFN.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_CN.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_CA.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_City.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_HM.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_B.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_FN.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Subscribe to GotFocus event for textBox_HM to automatically add "Dear ," when focused.
            textBox_HM.GotFocus += textBox_HM_GotFocus;

            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes (adjust as needed)
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Load previously saved settings.
            LoadUserSettings();
        }

        // Shown event handler: set watermark placeholders for single‑line textboxes.
        private void US_CoverLetter_Info_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // Instead of using WatermarkHelper, set properties directly.
                textBox_D.IsSingleLine = true;
                textBox_D.WatermarkText = " Date";

                textBox_HMFN.IsSingleLine = true;
                textBox_HMFN.WatermarkText = " John Smith";

                textBox_CN.IsSingleLine = true;
                textBox_CN.WatermarkText = " XYZ Solutions, Inc.";

                textBox_CA.IsSingleLine = true;
                textBox_CA.WatermarkText = " 1234 Madison Avenue";

                textBox_City.IsSingleLine = true;
                textBox_City.WatermarkText = " New York, NY 10110";

                textBox_HM.IsSingleLine = true;
                textBox_HM.WatermarkText = " Dear Mr. Smith,";

                textBox_FN.IsSingleLine = true;
                textBox_FN.WatermarkText = " Aiden Walsh";

                // Note: textBox_B is multiline, so we leave it unchanged.
            }
        }

        /// <summary>
        /// Helper method to determine design mode.
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
                // Optionally, update a status label here.
            }
        }

        // Saves cover letter data into user settings.
        private void AutosaveUserSettings()
        {
            Properties.Settings.Default.US_CoverLetter_Date = textBox_D.Text;
            Properties.Settings.Default.US_CoverLetter_HiringMgr = textBox_HM.Text;
            Properties.Settings.Default.US_CoverLetter_Company = textBox_CN.Text;
            Properties.Settings.Default.US_CoverLetter_CompAddress = textBox_CA.Text;
            Properties.Settings.Default.US_CoverLetter_City = textBox_City.Text;
            Properties.Settings.Default.US_CoverLetter_HMN = textBox_HMFN.Text;
            Properties.Settings.Default.US_CoverLetter_Body = textBox_B.Text;
            Properties.Settings.Default.US_CoverLetter_YFN = textBox_FN.Text;
            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        // Loads saved cover letter data into controls.
        private void LoadUserSettings()
        {
            textBox_D.Text = Properties.Settings.Default.US_CoverLetter_Date;
            textBox_HM.Text = Properties.Settings.Default.US_CoverLetter_HiringMgr;
            textBox_CN.Text = Properties.Settings.Default.US_CoverLetter_Company;
            textBox_CA.Text = Properties.Settings.Default.US_CoverLetter_CompAddress;
            textBox_City.Text = Properties.Settings.Default.US_CoverLetter_City;
            textBox_HMFN.Text = Properties.Settings.Default.US_CoverLetter_HMN;
            textBox_B.Text = Properties.Settings.Default.US_CoverLetter_Body;
            textBox_FN.Text = Properties.Settings.Default.US_CoverLetter_YFN;
        }

        private void US_CoverLetter_Info_FormClosing(object sender, FormClosingEventArgs e)
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

        // Navigation: Go to Cover Letter Template.
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
                // Data from CoverLetter (indicates 6 - 14)
                Properties.Settings.Default.US_CoverLetter_Date,
                Properties.Settings.Default.US_CoverLetter_HiringMgr,
                Properties.Settings.Default.US_CoverLetter_Company,
                Properties.Settings.Default.US_CoverLetter_CompAddress,
                Properties.Settings.Default.US_CoverLetter_City,
                Properties.Settings.Default.US_CoverLetter_HMN,
                Properties.Settings.Default.US_CoverLetter_Body,
                Properties.Settings.Default.US_CoverLetter_YFN
        };

            // 4) Launch the US cover letter form
            var coverLetterForm = new US_CoverLetter(userData)
            {
                Size = this.Size
            };
            coverLetterForm.Show();

            // 5) Close this form to free resources
            this.Close();
        }

        // Navigation: Go back to CoverLetter_Size.
        private void button_Back_Click(object sender, EventArgs e)
        {
            // 1) Flush any unsaved edits
            if (unsavedChanges)
                AutosaveUserSettings();
            // 2) Persist all Settings
            Properties.Settings.Default.Save();

            // 3) Launch the cover‑letter size chooser again
            var coverLetterSizeForm = new CoverLetter_Size
            {
                Size = this.Size
            };
            coverLetterSizeForm.Show();

            // 4) Close this form
            this.Close();
        }

        // Manual Save button event handler.
        private void button_Save_USCL_Click_1(object sender, EventArgs e)
        {
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Spell Check button event handler.
        private void button_SpellCheck_Click(object sender, EventArgs e)
        {
            try
            {
                // Initialize the spell checker and dictionary.
                Spelling spellChecker = new Spelling();
                WordDictionary dict = new WordDictionary();

                // Set the dictionary file path (ensure en-US.dic is in your project's output directory).
                string dictionaryPath = Path.Combine(Application.StartupPath, "Dictionaries", "en-US.dic");
                dict.DictionaryFile = dictionaryPath;
                dict.Initialize(); // Load the dictionary.
                spellChecker.Dictionary = dict;

                // Create a list to hold misspelled words.
                List<string> misspelledWords = new List<string>();

                // Define punctuation characters to trim.
                char[] punctuation = new char[] { ',', '.', ';', ':', '!', '?', '(', ')', '"', '\'' };

                // Regular expression to match words that consist solely of digits, dashes, and ampersands.
                Regex validPattern = new Regex(@"^[\d\-\&]+$");

                // Use TextBoxBase so both TextBox and WatermarkRichTextBox can be processed.
                TextBoxBase[] textBoxes = {
                    textBox_B
                };

                foreach (TextBoxBase tb in textBoxes)
                {
                    if (tb.ForeColor == Color.FromArgb(150, 150, 150))
                        continue;

                    string[] words = tb.Text
                        .Split(new[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string word in words)
                    {
                        string cleanWord = word.Trim(punctuation);
                        if (string.IsNullOrEmpty(cleanWord) || cleanWord.Contains("@") || validPattern.IsMatch(cleanWord))
                            continue;

                        try
                        {
                            if (!spellChecker.TestWord(cleanWord) &&
                                !misspelledWords.Contains(cleanWord))
                            {
                                misspelledWords.Add(cleanWord);
                            }
                        }
                        catch (IndexOutOfRangeException)
                        {
                            // NetSpell hiccuped on this particular token—just skip it
                        }
                    }
                }

                // Show the results in a MessageBox.
                if (misspelledWords.Count > 0)
                {
                    MessageBox.Show("The following words may be misspelled:\n" + string.Join(", ", misspelledWords),
                        "Spell Check Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    MessageBox.Show("No spelling errors found!", "Spell Check Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred during spell checking: " + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Event handler for when the hiring manager textbox (textBox_HM) gains focus.
        /// If the textbox is empty, it automatically inserts "Dear ," and positions the cursor
        /// between "Dear " and the comma. It also sets the text color to black and the background to white.
        /// </summary>
        private void textBox_HM_GotFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox_HM.Text))
            {
                textBox_HM.Text = "Dear ,";
                textBox_HM.SelectionStart = "Dear ".Length;
                textBox_HM.ForeColor = Color.Black;
                textBox_HM.BackColor = Color.White;
            }
        }
    }
}

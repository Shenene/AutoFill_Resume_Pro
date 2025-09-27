using AutoFill_Resume_Pro.Selection_Pages;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
// using AutoFill_Resume_Pro.Helpers; // Not needed anymore with the custom control
using System.IO;                         // For Path.Combine
using System.Text.RegularExpressions;    // For Regex
using System.Collections.Generic;        // For List<T>
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace AutoFill_Resume_Pro.Forms
{
    public partial class A4_Form_Skills : Form
    {
        // New fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        public string[] GetUserData()
        {
            return new string[]
            {
                textBoxPS_1.Text, // Professional Skill 1
                textBoxPS_2.Text, // Professional Skill 2
                textBoxPS_3.Text, // Professional Skill 3
                textBoxPS_4.Text, // Professional Skill 4
                textBoxPS_5.Text, // Professional Skill 5
                textBoxPS_6.Text, // Professional Skill 6
                textBoxPS_7.Text, // Professional Skill 7
                textBoxPS_8.Text, // Professional Skill 8
                textBoxPS_9.Text, // Professional Skill 9
                textBoxTS_1.Text, // Technical Skill 1
                textBoxTS_2.Text, // Technical Skill 2
                textBoxTS_3.Text, // Technical Skill 3
                textBoxTS_4.Text, // Technical Skill 4
                textBoxTS_5.Text, // Technical Skill 5
                textBoxTS_6.Text, // Technical Skill 6
                textBoxTS_7.Text, // Technical Skill 7
                textBoxTS_8.Text, // Technical Skill 8
                textBoxTS_9.Text, // Technical Skill 9
                };
        }

        public A4_Form_Skills()
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

            // Subscribe to Form Load, Shown, and FormClosing events.
            this.Load += A4_Form_Skills_Load;
            this.Shown += A4_Form_Skills_Shown;
            this.FormClosing += A4_Form_Skills_FormClosing;
        }

        public A4_Form_Skills(Size parentSize)
        {
            InitializeComponent();
            this.Size = parentSize;
            this.Load += A4_Form_Skills_Load;
            this.Shown += A4_Form_Skills_Shown;
            this.FormClosing += A4_Form_Skills_FormClosing;
        }

        private void A4_Form_Skills_Load(object sender, EventArgs e)
        {
            // Subscribe to TextChanged events for Professional Skills textboxes.
            textBoxPS_1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_6.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_7.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_8.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxPS_9.TextChanged += (s, ev) => { unsavedChanges = true; };
           
            // Subscribe to TextChanged events for Technical Skills textboxes.
            textBoxTS_1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_6.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_7.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_8.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBoxTS_9.TextChanged += (s, ev) => { unsavedChanges = true; };
           
            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes; adjust as needed.
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Load previously saved settings into the textboxes.
            LoadUserSettings();
        }

        /// <summary>
        /// Once the form is shown, set the watermark placeholders for all skill textboxes.
        /// </summary>
        private void A4_Form_Skills_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // Set placeholders for Professional Skills textboxes using custom control properties.
                textBoxPS_1.IsSingleLine = true;
                textBoxPS_1.WatermarkText = " Effective Communication";

                textBoxPS_2.IsSingleLine = true;
                textBoxPS_2.WatermarkText = " Team Collaboration";

                textBoxPS_3.IsSingleLine = true;
                textBoxPS_3.WatermarkText = " Problem-Solving Abilities";

                textBoxPS_4.IsSingleLine = true;
                textBoxPS_4.WatermarkText = " Time Management";

                textBoxPS_5.IsSingleLine = true;
                textBoxPS_5.WatermarkText = " Customer Service Excellence";

                textBoxPS_6.IsSingleLine = true;
                textBoxPS_6.WatermarkText = " Strategic Thinking";

                textBoxPS_7.IsSingleLine = true;
                textBoxPS_7.WatermarkText = " Adaptability and Flexibility";

                textBoxPS_8.IsSingleLine = true;
                textBoxPS_8.WatermarkText = " Interpersonal Skills";

                textBoxPS_9.IsSingleLine = true;
                textBoxPS_9.WatermarkText = " Attention to Detail";
                

                // Set placeholders for Technical Skills textboxes using custom control properties.
                textBoxTS_1.IsSingleLine = true;
                textBoxTS_1.WatermarkText = " Microsoft Office Suite";

                textBoxTS_2.IsSingleLine = true;
                textBoxTS_2.WatermarkText = " Email Management Systems";

                textBoxTS_3.IsSingleLine = true;
                textBoxTS_3.WatermarkText = " Basic Data Entry and Analysis";

                textBoxTS_4.IsSingleLine = true;
                textBoxTS_4.WatermarkText = " Internet Research Skills";

                textBoxTS_5.IsSingleLine = true;
                textBoxTS_5.WatermarkText = " Basic Graphic Design";

                textBoxTS_6.IsSingleLine = true;
                textBoxTS_6.WatermarkText = " Database Management";

                textBoxTS_7.IsSingleLine = true;
                textBoxTS_7.WatermarkText = " Enterprise Resource Planning";

                textBoxTS_8.IsSingleLine = true;
                textBoxTS_8.WatermarkText = " Social Media Management";

                textBoxTS_9.IsSingleLine = true;
                textBoxTS_9.WatermarkText = " Video Conferencing Tools";                
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

        // Save the Professional and Technical skills.
        private void AutosaveUserSettings()
        {
            // Professional Skills
            Properties.Settings.Default.A4Skills_ProSkill_1 = textBoxPS_1.Text;
            Properties.Settings.Default["A4Skills_ProSkill_2"] = textBoxPS_2.Text;
            Properties.Settings.Default["A4Skills_ProSkill_3"] = textBoxPS_3.Text;
            Properties.Settings.Default["A4Skills_ProSkill_4"] = textBoxPS_4.Text;
            Properties.Settings.Default["A4Skills_ProSkill_5"] = textBoxPS_5.Text;
            Properties.Settings.Default["A4Skills_ProSkill_6"] = textBoxPS_6.Text;
            Properties.Settings.Default["A4Skills_ProSkill_7"] = textBoxPS_7.Text;
            Properties.Settings.Default["A4Skills_ProSkill_8"] = textBoxPS_8.Text;
            Properties.Settings.Default["A4Skills_ProSkill_9"] = textBoxPS_9.Text;           

            // Technical Skills
            Properties.Settings.Default.A4Skills_TechSkill_1 = textBoxTS_1.Text;
            Properties.Settings.Default["A4Skills_TechSkill_2"] = textBoxTS_2.Text;
            Properties.Settings.Default["A4Skills_TechSkill_3"] = textBoxTS_3.Text;
            Properties.Settings.Default["A4Skills_TechSkill_4"] = textBoxTS_4.Text;
            Properties.Settings.Default["A4Skills_TechSkill_5"] = textBoxTS_5.Text;
            Properties.Settings.Default["A4Skills_TechSkill_6"] = textBoxTS_6.Text;
            Properties.Settings.Default["A4Skills_TechSkill_7"] = textBoxTS_7.Text;
            Properties.Settings.Default["A4Skills_TechSkill_8"] = textBoxTS_8.Text;
            Properties.Settings.Default["A4Skills_TechSkill_9"] = textBoxTS_9.Text;           

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        // Load saved settings into the textboxes.
        private void LoadUserSettings()
        {
            // Professional Skills
            textBoxPS_1.Text = Properties.Settings.Default.A4Skills_ProSkill_1;
            textBoxPS_2.Text = Properties.Settings.Default["A4Skills_ProSkill_2"].ToString();
            textBoxPS_3.Text = Properties.Settings.Default["A4Skills_ProSkill_3"].ToString();
            textBoxPS_4.Text = Properties.Settings.Default["A4Skills_ProSkill_4"].ToString();
            textBoxPS_5.Text = Properties.Settings.Default["A4Skills_ProSkill_5"].ToString();
            textBoxPS_6.Text = Properties.Settings.Default["A4Skills_ProSkill_6"].ToString();
            textBoxPS_7.Text = Properties.Settings.Default["A4Skills_ProSkill_7"].ToString();
            textBoxPS_8.Text = Properties.Settings.Default["A4Skills_ProSkill_8"].ToString();
            textBoxPS_9.Text = Properties.Settings.Default["A4Skills_ProSkill_9"].ToString();            

            // Technical Skills
            textBoxTS_1.Text = Properties.Settings.Default.A4Skills_TechSkill_1;
            textBoxTS_2.Text = Properties.Settings.Default["A4Skills_TechSkill_2"].ToString();
            textBoxTS_3.Text = Properties.Settings.Default["A4Skills_TechSkill_3"].ToString();
            textBoxTS_4.Text = Properties.Settings.Default["A4Skills_TechSkill_4"].ToString();
            textBoxTS_5.Text = Properties.Settings.Default["A4Skills_TechSkill_5"].ToString();
            textBoxTS_6.Text = Properties.Settings.Default["A4Skills_TechSkill_6"].ToString();
            textBoxTS_7.Text = Properties.Settings.Default["A4Skills_TechSkill_7"].ToString();
            textBoxTS_8.Text = Properties.Settings.Default["A4Skills_TechSkill_8"].ToString();
            textBoxTS_9.Text = Properties.Settings.Default["A4Skills_TechSkill_9"].ToString();           
        }

        private void A4_Form_Skills_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Autosave any unsaved changes when closing.
            if (unsavedChanges)
            {
                AutosaveUserSettings();
            }
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

        // Navigation: Go to A4_Form_WorkExperience.
        private void button_Next_Click(object sender, EventArgs e)
        {
            // 1) Flush any edits into Settings
            if (unsavedChanges)
                AutosaveUserSettings();
            Properties.Settings.Default.Save();

            // 2) Open WorkExperience and close this form
            var workExp = new A4_Form_WorkExperience { Size = this.Size };
            workExp.Show();
            this.Close();
        }

        // Navigation: Go back to A4_Form_ProfessionalInfo.
        private void button_Back_Click(object sender, EventArgs e)
        {
            // 1) Flush any edits into Settings
            if (unsavedChanges)
                AutosaveUserSettings();
            Properties.Settings.Default.Save();

            // 2) Open ProfessionalInfo and close this form
            var profInfo = new A4_Form_ProfessionalInfo { Size = this.Size };
            profInfo.Show();
            this.Close();
        }

        // Save button event handler for A4 Skills.
        private void button_Save_A4Skills_Click_1(object sender, EventArgs e)
        {
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Spell Check button event handler.
        private void button_SpellCheck_Click(object sender, EventArgs e)
        {
            try
            {
                // Initialize the spell checker and dictionary
                Spelling spellChecker = new Spelling();
                WordDictionary dict = new WordDictionary();

                // Set the dictionary file path (ensure en-US.dic is in your project's output directory)
                string dictionaryPath = Path.Combine(Application.StartupPath, "Dictionaries", "en-US.dic");
                dict.DictionaryFile = dictionaryPath;
                dict.Initialize(); // Load the dictionary
                spellChecker.Dictionary = dict;

                // Create a list to hold misspelled words
                List<string> misspelledWords = new List<string>();

                // Define punctuation characters to trim
                char[] punctuation = new char[] { ',', '.', ';', ':', '!', '?', '(', ')', '"', '\'' };

                // Regular expression to match words that consist solely of digits, dashes, and ampersands
                Regex validPattern = new Regex(@"^[\d\-\&]+$");

                // Spell check all relevant textboxes on the form (both Professional and Technical skills)
                TextBoxBase[] textBoxes = {
                    textBoxPS_1, textBoxPS_2, textBoxPS_3, textBoxPS_4, textBoxPS_5, textBoxPS_6, textBoxPS_7, textBoxPS_8,
                    textBoxPS_9, 
                    textBoxTS_1, textBoxTS_2, textBoxTS_3, textBoxTS_4, textBoxTS_5, textBoxTS_6, textBoxTS_7, textBoxTS_8,
                    textBoxTS_9
                };

                foreach (TextBoxBase tb in textBoxes)
                {
                    // Optionally, skip placeholder text by checking the ForeColor (adjust if needed)
                    if (tb.ForeColor == Color.FromArgb(150, 150, 150))
                        continue;

                    // Split the text into words (simple split, adjust if needed)
                    string[] words = tb.Text.Split(new char[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string word in words)
                    {
                        // Clean up the word by trimming punctuation
                        string cleanWord = word.Trim(punctuation);

                        // Only check if the cleaned word is not empty
                        if (string.IsNullOrEmpty(cleanWord))
                            continue;

                        // Skip words that appear to be email addresses
                        if (cleanWord.Contains("@"))
                            continue;

                        // If the word matches our valid pattern (only digits, dashes, and ampersands), skip checking
                        if (validPattern.IsMatch(cleanWord))
                            continue;

                        // Check if the word is spelled correctly
                        if (!spellChecker.TestWord(cleanWord))
                        {
                            if (!misspelledWords.Contains(cleanWord))
                                misspelledWords.Add(cleanWord);
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
    }
}

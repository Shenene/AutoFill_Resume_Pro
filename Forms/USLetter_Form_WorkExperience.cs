using AutoFill_Resume_Pro.Templates;
using AutoFill_Resume_Pro.Controls; // For WatermarkRichTextBox
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace AutoFill_Resume_Pro.Forms
{
    public partial class USLetter_Form_WorkExperience : Form
    {
        // Fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        public string[] GetUserData()
        {
            return new string[]
            {
                textBox_Position1.Text, // Position 1
                textBox_Company1.Text,  // Company 1
                textBox_Location1.Text, // Location 1
                textBox_Year1.Text,     // Year 1
                textBox_Exp1.Text,      // Experience 1
                textBox_Position2.Text, // Position 2
                textBox_Company2.Text,  // Company 2
                textBox_Location2.Text, // Location 2
                textBox_Year2.Text,     // Year 2
                textBox_Exp2.Text       // Experience 2
            };
        }

        public USLetter_Form_WorkExperience()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            this.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Regular, GraphicsUnit.Point);
            InitializeComponent();

            // For multiline textboxes, set their WatermarkText.
            textBox_Exp1.WatermarkText = @"• Bullet Points: Describe your job role in 2-3 bullet points. Avoid using too many Bullets.
• Quantify Results: Whenever possible, use numbers to showcase your achievements (e.g., increased sales by 20%, managed a team of 15).
• Highlight Responsibilities: Clearly outline your key responsibilities and how they contributed to the success of the organization.
• Demonstrate Growth: Include examples of how you've grown in your roles, such as taking on additional responsibilities or earning promotions.
• Focus on Impact: Describe how your contributions made a positive impact on the company, team, or project.
• Use Action Verbs: Start each bullet point with a strong action verb to convey your contributions clearly (e.g., led, developed, implemented).";

            textBox_Exp2.WatermarkText = @"• Bullet Points: Describe your job role in 2-3 bullet points. Avoid using too many Bullets.
• Quantify Results: Whenever possible, use numbers to showcase your achievements (e.g., increased sales by 20%, managed a team of 15).
• Highlight Responsibilities: Clearly outline your key responsibilities and how they contributed to the success of the organization.
• Demonstrate Growth: Include examples of how you've grown in your roles, such as taking on additional responsibilities or earning promotions.
• Focus on Impact: Describe how your contributions made a positive impact on the company, team, or project.
• Use Action Verbs: Start each bullet point with a strong action verb to convey your contributions clearly (e.g., led, developed, implemented).";

            this.ClientSize = new Size(800, 950);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(800, 768);
            this.MaximizeBox = false;
            this.DoubleBuffered = true;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 874);

            this.Paint += Resume_Size_Paint;

            // Subscribe to Form Load, Shown, and FormClosing events.
            this.Load += USLetter_Form_WorkExperience_Load;
            this.Shown += USLetter_Form_WorkExperience_Shown;
            this.FormClosing += USLetter_Form_WorkExperience_FormClosing;
        }

        public USLetter_Form_WorkExperience(Size parentSize)
        {
            InitializeComponent();
            // Set watermark for multiline textboxes.
            textBox_Exp1.WatermarkText = @"• Bullet Points: Describe your job role in 2-3 bullet points. Avoid using too many Bullets.
• Quantify Results: Whenever possible, use numbers to showcase your achievements (e.g., increased sales by 20%, managed a team of 15).
• Highlight Responsibilities: Clearly outline your key responsibilities and how they contributed to the success of the organization.
• Demonstrate Growth: Include examples of how you've grown in your roles, such as taking on additional responsibilities or earning promotions.
• Focus on Impact: Describe how your contributions made a positive impact on the company, team, or project.
• Use Action Verbs: Start each bullet point with a strong action verb to convey your contributions clearly (e.g., led, developed, implemented).";

            textBox_Exp2.WatermarkText = @"• Bullet Points: Describe your job role in 2-3 bullet points. Avoid using too many Bullets.
• Quantify Results: Whenever possible, use numbers to showcase your achievements (e.g., increased sales by 20%, managed a team of 15).
• Highlight Responsibilities: Clearly outline your key responsibilities and how they contributed to the success of the organization.
• Demonstrate Growth: Include examples of how you've grown in your roles, such as taking on additional responsibilities or earning promotions.
• Focus on Impact: Describe how your contributions made a positive impact on the company, team, or project.
• Use Action Verbs: Start each bullet point with a strong action verb to convey your contributions clearly (e.g., led, developed, implemented).";

            this.Size = parentSize;
            this.Load += USLetter_Form_WorkExperience_Load;
            this.Shown += USLetter_Form_WorkExperience_Shown;
            this.FormClosing += USLetter_Form_WorkExperience_FormClosing;
        }

        private void USLetter_Form_WorkExperience_Load(object sender, EventArgs e)
        {
            // Subscribe to TextChanged events for autosave detection.
            textBox_Position1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Company1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Location1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Year1.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Exp1.TextChanged += (s, ev) => { unsavedChanges = true; };

            textBox_Position2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Company2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Location2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Year2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox_Exp2.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes; adjust as needed.
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Load previously saved settings.
            LoadUserSettings();

            // Subscribe to GotFocus and KeyDown events for automatic bullet points in experience textboxes.
            textBox_Exp1.GotFocus += TextBox_Exp_GotFocus;
            textBox_Exp2.GotFocus += TextBox_Exp_GotFocus;
            textBox_Exp1.KeyDown += TextBox_Exp_KeyDown;
            textBox_Exp2.KeyDown += TextBox_Exp_KeyDown;
        }

        // Shown event handler: set watermark placeholders for single‑line textboxes.
        private void USLetter_Form_WorkExperience_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // Instead of using WatermarkHelper, set properties directly.
                textBox_Position1.IsSingleLine = true;
                textBox_Position1.WatermarkText = " POSITION TITLE HERE";

                textBox_Company1.IsSingleLine = true;
                textBox_Company1.WatermarkText = " Company";

                textBox_Location1.IsSingleLine = true;
                textBox_Location1.WatermarkText = " City, State";

                textBox_Year1.IsSingleLine = true;
                textBox_Year1.WatermarkText = " Year - Present";

                textBox_Position2.IsSingleLine = true;
                textBox_Position2.WatermarkText = " POSITION TITLE HERE";

                textBox_Company2.IsSingleLine = true;
                textBox_Company2.WatermarkText = " Company";

                textBox_Location2.IsSingleLine = true;
                textBox_Location2.WatermarkText = " City, State";

                textBox_Year2.IsSingleLine = true;
                textBox_Year2.WatermarkText = " Year - Year";
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

        // Saves the current work experience values into user settings.
        private void AutosaveUserSettings()
        {
            Properties.Settings.Default.US_WorkExp_Position1 = textBox_Position1.Text;
            Properties.Settings.Default.US_WorkExp_Company1 = textBox_Company1.Text;
            Properties.Settings.Default.US_WorkExp_Location1 = textBox_Location1.Text;
            Properties.Settings.Default.US_WorkExp_Year1 = textBox_Year1.Text;
            Properties.Settings.Default.US_WorkExp_Experience1 = textBox_Exp1.Text;

            Properties.Settings.Default.US_WorkExp_Position2 = textBox_Position2.Text;
            Properties.Settings.Default.US_WorkExp_Company2 = textBox_Company2.Text;
            Properties.Settings.Default.US_WorkExp_Location2 = textBox_Location2.Text;
            Properties.Settings.Default.US_WorkExp_Year2 = textBox_Year2.Text;
            Properties.Settings.Default.US_WorkExp_Experience2 = textBox_Exp2.Text;

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        // Loads saved settings into the textboxes.
        private void LoadUserSettings()
        {
            textBox_Position1.Text = Properties.Settings.Default.US_WorkExp_Position1;
            textBox_Company1.Text = Properties.Settings.Default.US_WorkExp_Company1;
            textBox_Location1.Text = Properties.Settings.Default.US_WorkExp_Location1;
            textBox_Year1.Text = Properties.Settings.Default.US_WorkExp_Year1;
            textBox_Exp1.Text = Properties.Settings.Default.US_WorkExp_Experience1;

            textBox_Position2.Text = Properties.Settings.Default.US_WorkExp_Position2;
            textBox_Company2.Text = Properties.Settings.Default.US_WorkExp_Company2;
            textBox_Location2.Text = Properties.Settings.Default.US_WorkExp_Location2;
            textBox_Year2.Text = Properties.Settings.Default.US_WorkExp_Year2;
            textBox_Exp2.Text = Properties.Settings.Default.US_WorkExp_Experience2;
        }

        private void USLetter_Form_WorkExperience_FormClosing(object sender, FormClosingEventArgs e)
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

        // Navigation: Go to USLetter_Resume (Resume Template).
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
                // Data from ProfessionalInfo (indicates 6 - 14)
                Properties.Settings.Default.USProInfo_ProSum,
                Properties.Settings.Default.USProInfo_Cert_1,
                Properties.Settings.Default.USProInfo_Major_1,
                Properties.Settings.Default.USProInfo_College_1,
                Properties.Settings.Default.USProInfo_Year_1,
                Properties.Settings.Default.USProInfo_Cert_2,
                Properties.Settings.Default.USProInfo_Major_2,
                Properties.Settings.Default.USProInfo_College_2,
                Properties.Settings.Default.USProInfo_Year_2,
                // Data from Skills (indicates 15 - 32)
                Properties.Settings.Default.USSkills_ProSkill_1,
                Properties.Settings.Default.USSkills_ProSkill_2,
                Properties.Settings.Default.USSkills_ProSkill_3,
                Properties.Settings.Default.USSkills_ProSkill_4,
                Properties.Settings.Default.USSkills_ProSkill_5,
                Properties.Settings.Default.USSkills_ProSkill_6,
                Properties.Settings.Default.USSkills_ProSkill_7,
                Properties.Settings.Default.USSkills_ProSkill_8,
                Properties.Settings.Default.USSkills_ProSkill_9,
                Properties.Settings.Default.USSkills_TechSkill_1,
                Properties.Settings.Default.USSkills_TechSkill_2,
                Properties.Settings.Default.USSkills_TechSkill_3,
                Properties.Settings.Default.USSkills_TechSkill_4,
                Properties.Settings.Default.USSkills_TechSkill_5,
                Properties.Settings.Default.USSkills_TechSkill_6,
                Properties.Settings.Default.USSkills_TechSkill_7,
                Properties.Settings.Default.USSkills_TechSkill_8,
                Properties.Settings.Default.USSkills_TechSkill_9,
                // Data from WorkExperience (indicates 33 - 42)
                Properties.Settings.Default.US_WorkExp_Position1,
                Properties.Settings.Default.US_WorkExp_Company1,
                Properties.Settings.Default.US_WorkExp_Location1,
                Properties.Settings.Default.US_WorkExp_Year1,
                Properties.Settings.Default.US_WorkExp_Experience1,
                Properties.Settings.Default.US_WorkExp_Position2,
                Properties.Settings.Default.US_WorkExp_Company2,
                Properties.Settings.Default.US_WorkExp_Location2,
                Properties.Settings.Default.US_WorkExp_Year2,
                Properties.Settings.Default.US_WorkExp_Experience2
            };

            var usResume = new USLetter_Resume(userData)
            {
                Size = this.Size
            };
            usResume.Show();

            // 4) Close this form
            this.Close();
        }

        // Navigation: Go back to USLetter_Form_Skills.
        private void button_Back_Click(object sender, EventArgs e)
        {
            // 1) Flush any unsaved edits into Settings
            if (unsavedChanges)
                AutosaveUserSettings();
            // 2) Persist to disk
            Properties.Settings.Default.Save();

            // 3) Open the Skills form
            var skillsForm = new USLetter_Form_Skills
            {
                Size = this.Size
            };
            skillsForm.Show();

            // 4) Close this form
            this.Close();
        }

        // Save button event handler for Work Experience.
        private void button_Save_USWork_Click_1(object sender, EventArgs e)
        {
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Spell Check button event handler.
        private void button_SpellCheck_Click_1(object sender, EventArgs e)
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

                // Define punctuation to trim (added • and %)
                char[] punctuation = { ',', '.', ';', ':', '!', '?', '(', ')', '"', '\'', '•', '%' };

                // Abbreviations or symbols to skip outright (note both "e.g" and "e.g.")
                var skipWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
{
    "e.g", "e.g.", "i.e", "i.e."
};


                // Regular expression to match words that consist solely of digits, dashes, and ampersands.
                Regex validPattern = new Regex(@"^[\d\-\&]+$");

                // Use TextBoxBase so both TextBox and WatermarkRichTextBox can be processed.
                TextBoxBase[] textBoxes = {
                    textBox_Position1, textBox_Location1, textBox_Exp1,
                    textBox_Position2, textBox_Location2, textBox_Exp2
                };

                foreach (TextBoxBase tb in textBoxes)
                {
                    // Optionally, skip placeholder text by checking the ForeColor (adjust if needed).
                    if (tb.ForeColor == Color.FromArgb(150, 150, 150))
                        continue;

                    // Split the text into words (simple split, adjust if needed).
                    string[] words = tb.Text.Split(new char[] { ' ', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var word in words)
                    {
                        // 1) Trim any leading/trailing punctuation (• , % etc)
                        string cleanWord = word.Trim(punctuation);

                        if (string.IsNullOrEmpty(cleanWord))
                            continue;

                        // Skip words that appear to be email addresses.
                        if (cleanWord.Contains("@"))
                            continue;

                        // 3) Skip our hard‑coded abbreviations
                        if (skipWords.Contains(cleanWord))
                            continue;

                        // If the word matches our valid pattern (only digits, dashes, and ampersands), skip checking.
                        if (validPattern.IsMatch(cleanWord))
                            continue;

                        // Check if the word is spelled correctly.
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

        /// <summary>
        /// Event handler for when the experience textboxes gain focus.
        /// If the textbox is empty or missing the bullet at the start, it automatically inserts "• " 
        /// and sets the caret right after the bullet.
        /// </summary>
        private void TextBox_Exp_GotFocus(object sender, EventArgs e)
        {
            TextBoxBase txt = sender as TextBoxBase;
            if (txt != null)
            {
                if (string.IsNullOrWhiteSpace(txt.Text))
                {
                    txt.Text = "• ";
                    this.BeginInvoke(new Action(() => { txt.SelectionStart = txt.Text.Length; }));
                }
                else if (!txt.Text.StartsWith("• "))
                {
                    txt.Text = "• " + txt.Text;
                    this.BeginInvoke(new Action(() => { txt.SelectionStart = txt.Text.Length; }));
                }
            }
        }

        /// <summary>
        /// Event handler for KeyDown in the experience textboxes.
        /// Inserts a newline and bullet ("• ") automatically when the Enter key is pressed.
        /// </summary>
        private void TextBox_Exp_KeyDown(object sender, KeyEventArgs e)
        {
            TextBoxBase txt = sender as TextBoxBase;
            if (txt != null && e.KeyCode == Keys.Enter)
            {
                int selPos = txt.SelectionStart;
                string bulletText = "\n• ";
                txt.Text = txt.Text.Insert(selPos, bulletText);
                txt.SelectionStart = selPos + bulletText.Length;
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }
    }
}

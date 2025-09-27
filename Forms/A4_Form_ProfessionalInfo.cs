using AutoFill_Resume_Pro.Selection_Pages;
using AutoFill_Resume_Pro.Controls;  // Reference the custom control (WatermarkRichTextBox)
using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;                         // For Path.Combine
using System.Text.RegularExpressions;    // For Regex
using System.Collections.Generic;        // For List<T>
using NetSpell.SpellChecker;
using NetSpell.SpellChecker.Dictionary;

namespace AutoFill_Resume_Pro.Forms
{
    public partial class A4_Form_ProfessionalInfo : Form
    {
        // New fields for autosave and unsaved changes detection.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        public string[] GetUserData()
        {
            return new string[]
            {
                textBox_ProSum.Text, // Pro Summary
                textBox2.Text, // Cert 1
                textBox3.Text, // Major 1
                textBox4.Text, // College 1
                textBox5.Text, // Year 1
                textBox6.Text, // Cert 2
                textBox7.Text, // Major 2
                textBox8.Text, // College 2
                textBox9.Text  // Year 2
            };
        }

        public A4_Form_ProfessionalInfo()
        {
            this.AutoScaleMode = AutoScaleMode.None;
            this.Font = new Font(this.Font.FontFamily, 10F, FontStyle.Regular, GraphicsUnit.Point);
            InitializeComponent();

            // Ensure textBox1 is set to multiline mode.
            textBox_ProSum.IsSingleLine = false;
            // Set the WatermarkText for the multiline textbox.
            textBox_ProSum.WatermarkText = "Use this section to make a strong first impression by summarizing your key skills, experiences, and career highlights. Focus on what makes you an exceptional candidate for a wide range of roles, showcasing your versatility and ability to adapt to different industries. Highlight your core competencies and notable achievements to capture the reader's attention and set the tone for the rest of your resume.";

            this.ClientSize = new Size(800, 950);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(800, 768);
            this.MaximizeBox = false;
            this.DoubleBuffered = true;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 874);

            this.Paint += Resume_Size_Paint;

            // Subscribe to Form Load, Shown, and FormClosing events.
            this.Load += A4_Form_ProfessionalInfo_Load;
            this.Shown += A4_Form_ProfessionalInfo_Shown;
            this.FormClosing += A4_Form_ProfessionalInfo_FormClosing;
        }

        public A4_Form_ProfessionalInfo(Size parentSize)
        {
            InitializeComponent();
            // Ensure textBox1 is in multiline mode.
            textBox_ProSum.IsSingleLine = false;
            textBox_ProSum.WatermarkText = "Use this section to make a strong first impression by summarizing your key skills, experiences, and career highlights. Focus on what makes you an exceptional candidate for a wide range of roles, showcasing your versatility and ability to adapt to different industries. Highlight your core competencies and notable achievements to capture the reader's attention and set the tone for the rest of your resume.";

            this.Size = parentSize;
            // Subscribe to Form Load, Shown, and FormClosing events.
            this.Load += A4_Form_ProfessionalInfo_Load;
            this.Shown += A4_Form_ProfessionalInfo_Shown;
            this.FormClosing += A4_Form_ProfessionalInfo_FormClosing;
        }

        private void A4_Form_ProfessionalInfo_Load(object sender, EventArgs e)
        {
            // Subscribe to TextChanged events for autosave detection.
            textBox_ProSum.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox2.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox3.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox4.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox5.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox6.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox7.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox8.TextChanged += (s, ev) => { unsavedChanges = true; };
            textBox9.TextChanged += (s, ev) => { unsavedChanges = true; };

            // Initialize autosave timer.
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes; adjust as needed.
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();

            // Load previously saved settings.
            LoadUserSettings();
        }

        /// <summary>
        /// Once the form is shown, set the watermark placeholders for single-line text boxes.
        /// For textBox1 (multiline), its WatermarkText property is already set.
        /// </summary>
        private void A4_Form_ProfessionalInfo_Shown(object sender, EventArgs e)
        {
            if (!IsInDesignMode())
            {
                // For textBox2 to textBox9, set the WatermarkText property.
                textBox2.IsSingleLine = true;
                textBox2.WatermarkText = " Degree / Diploma / Certificate";

                textBox3.IsSingleLine = true;
                textBox3.WatermarkText = " Major";

                textBox4.IsSingleLine = true;
                textBox4.WatermarkText = " College, University";

                textBox5.IsSingleLine = true;
                textBox5.WatermarkText = " Year - Year";

                textBox6.IsSingleLine = true;
                textBox6.WatermarkText = " Degree / Diploma / Certificate";

                textBox7.IsSingleLine = true;
                textBox7.WatermarkText = " Major";

                textBox8.IsSingleLine = true;
                textBox8.WatermarkText = " College, University";

                textBox9.IsSingleLine = true;
                textBox9.WatermarkText = " Year - Year";
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

        // Save the user settings for A4 Professional Info.
        private void AutosaveUserSettings()
        {
            Properties.Settings.Default.A4ProInfo_ProSum = textBox_ProSum.Text;
            Properties.Settings.Default.A4ProInfo_Cert_1 = textBox2.Text;
            Properties.Settings.Default.A4ProInfo_Major_1 = textBox3.Text;
            Properties.Settings.Default.A4ProInfo_College_1 = textBox4.Text;
            Properties.Settings.Default.A4ProInfo_Year_1 = textBox5.Text;
            Properties.Settings.Default.A4ProInfo_Cert_2 = textBox6.Text;
            Properties.Settings.Default.A4ProInfo_Major_2 = textBox7.Text;
            Properties.Settings.Default.A4ProInfo_College_2 = textBox8.Text;
            Properties.Settings.Default.A4ProInfo_Year_2 = textBox9.Text;
            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        // Load saved settings back into the textboxes.
        private void LoadUserSettings()
        {
            textBox_ProSum.Text = Properties.Settings.Default.A4ProInfo_ProSum;
            textBox2.Text = Properties.Settings.Default.A4ProInfo_Cert_1;
            textBox3.Text = Properties.Settings.Default.A4ProInfo_Major_1;
            textBox4.Text = Properties.Settings.Default.A4ProInfo_College_1;
            textBox5.Text = Properties.Settings.Default.A4ProInfo_Year_1;
            textBox6.Text = Properties.Settings.Default.A4ProInfo_Cert_2;
            textBox7.Text = Properties.Settings.Default.A4ProInfo_Major_2;
            textBox8.Text = Properties.Settings.Default.A4ProInfo_College_2;
            textBox9.Text = Properties.Settings.Default.A4ProInfo_Year_2;
        }

        private void A4_Form_ProfessionalInfo_FormClosing(object sender, FormClosingEventArgs e)
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
        private void button_Return_Click(object sender, EventArgs e)
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

        // Navigation: Go to A4_Form_Skills.
        private void button_Next_Click(object sender, EventArgs e)
        {
            // 1) Flush any edits into Settings
            if (unsavedChanges)
                AutosaveUserSettings();
            Properties.Settings.Default.Save();

            // 2) Open Skills form and close this one
            var skillsForm = new A4_Form_Skills { Size = this.Size };
            skillsForm.Show();
            this.Close();
        }

        // Navigation: Go back to Resume_Size.
        private void button_Back_Click(object sender, EventArgs e)
        {
            // 1) Flush any edits into Settings
            if (unsavedChanges)
                AutosaveUserSettings();
            Properties.Settings.Default.Save();

            // 2) Open Resume Size form and close this one
            var resumeSize = new Resume_Size { Size = this.Size };
            resumeSize.Show();
            this.Close();
        }

        // Save button event handler for A4 Professional Info.
        private void button_Save_A4ProInfo_Click_1(object sender, EventArgs e)
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

                // Use TextBoxBase so both TextBox and WatermarkRichTextBox can be processed.
                TextBoxBase[] textBoxes = { textBox_ProSum, textBox2, textBox3, (TextBoxBase)textBox4, textBox6, textBox7, textBox8 };

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
                MessageBox.Show("An error occurred during spell checking: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}

// Place WatermarkHelper after the form class so that the designer sees the form class first.
namespace AutoFill_Resume_Pro.Forms
{
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

using AutoFill_Resume_Pro.Forms;
using AutoFill_Resume_Pro.Selection_Pages;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using static System.Windows.Forms.AxHost;

namespace AutoFill_Resume_Pro.Templates
{
    public partial class A4_Resume : Form
    {
        // The fixed A4 print dimensions (at 96 DPI) and on‑screen design dimensions.
        private const int fixedWidth = 794;
        private const int fixedHeight = 1123;

        private PrintDocument printDocument;
        // We still keep originalControlBounds if you need to adjust positions manually later.
        private Dictionary<Control, Rectangle> originalControlBounds = new Dictionary<Control, Rectangle>();

        private Image defaultIcon1;
        private Image defaultIcon2;
        private Image defaultIcon3;
        private Image defaultIcon4;        

        // Autosave variables
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        // Constructor that accepts userData.
        public A4_Resume(string[] userData)
        {
            // 1) Base sizing & scrolling
            this.AutoScaleMode = AutoScaleMode.None;
            this.ClientSize = new Size(810, 1350);
            this.MinimumSize = new Size(810, 768);
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 950);

            InitializeComponent();

            panelA4.BackColor = Color.White;

            this.Load += A4_Resume_Load;

            // 2) Hook Load exactly once, and do everything there:
            this.Load += (s, e) =>
            {
                // a) Load saved fonts/colors/icons
                LoadUserSettings();

                // b) Now override with userData (must be at least 43 items)
                if (userData != null && userData.Length >= 43)
                {
                    // Header
                    label_Name.Text = userData[0];
                    label_PositionTitle.Text = userData[1];
                    label_Phone.Text = userData[2];
                    label_Email.Text = userData[3];
                    label_Location.Text = userData[4];
                    label_LinkedIn.Text = userData[5];

                    // Professional summary & education (6–14)
                    label_ProfessionalSummary.Text = userData[6];
                    label_Degree_1.Text = userData[7];
                    label_Major_1.Text = userData[8];
                    label_College_1.Text = userData[9];
                    label_Year_1.Text = userData[10];
                    label_Degree_2.Text = userData[11];
                    label_Major_2.Text = userData[12];
                    label_College_2.Text = userData[13];
                    label_Year_2.Text = userData[14];

                    // Skills (15–32)
                    label_Skill_PS1.Text = userData[15];
                    label_Skill_PS2.Text = userData[16];
                    label_Skill_PS3.Text = userData[17];
                    label_Skill_PS4.Text = userData[18];
                    label_Skill_PS5.Text = userData[19];
                    label_Skill_PS6.Text = userData[20];
                    label_Skill_PS7.Text = userData[21];
                    label_Skill_PS8.Text = userData[22];
                    label_Skill_PS9.Text = userData[23];
                    label_Skill_TS1.Text = userData[24];
                    label_Skill_TS2.Text = userData[25];
                    label_Skill_TS3.Text = userData[26];
                    label_Skill_TS4.Text = userData[27];
                    label_Skill_TS5.Text = userData[28];
                    label_Skill_TS6.Text = userData[29];
                    label_Skill_TS7.Text = userData[30];
                    label_Skill_TS8.Text = userData[31];
                    label_Skill_TS9.Text = userData[32];

                    // Work Experience (33–42)
                    label_Position_1.Text = userData[33];
                    label_Company_1.Text = userData[34];
                    label_City_1.Text = userData[35];
                    label_Y1.Text = userData[36];
                    label_Work_1.Text = userData[37];

                    label_Position_2.Text = userData[38];
                    label_Company_2.Text = userData[39];
                    label_City_2.Text = userData[40];
                    label_Y2.Text = userData[41];
                    label_Work_2.Text = userData[42];
                }
            };        

        // Initialize the print document.
        printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

            // Set the default page settings to A4.
            // A4 is approximately 8.27" x 11.69", which translates to 827 x 1169 (in hundredths of an inch).
            PaperSize a4Size = new PaperSize("A4", 827, 1169);
            printDocument.DefaultPageSettings.PaperSize = a4Size;
            // Set margins to zero since we'll draw our own white background.
            printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            // Recenter panelA4 on resize.
            this.Resize += A4_Resume_Resize;

            // Load fonts into all ComboBoxes.
            LoadFonts();

            // Set default font size and color for all sections.
            numericUpDown_Size_Name.Value = 25; // Name
            numericUpDown_Size_PT.Value = 12;   // Position Title
            numericUpDown_Size_H.Value = 12;    // Headings
            numericUpDown_Size_S.Value = 10;    // Subheadings
            numericUpDown_Size_B.Value = 10;    // Body
            button_Color_Name.BackColor = Color.Black;
            button_Color_PT.BackColor = Color.Black;
            button_Color_H.BackColor = Color.Black;
            button_Color_S.BackColor = Color.Black;
            button_Color_B.BackColor = Color.Black;
            // Set default for the separate icon color button.
            button_IconColor.BackColor = Color.Black;

            // Attach event handlers for the Name section controls.
            comboBox_FontName_Name.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_Name.ValueChanged += FontSettingsChanged;
            button_Color_Name.Click += btnFontColor_Click;

            // Attach event handlers for the Position Title section.
            comboBox_FontName_PT.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_PT.ValueChanged += FontSettingsChanged;
            button_Color_PT.Click += btnColor_PT_Click;

            // Attach event handlers for the Heading section.
            comboBox_FontName_H.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_H.ValueChanged += FontSettingsChanged;
            button_Color_H.Click += btnColor_H_Click;

            // Attach event handlers for the Subheading section.
            comboBox_FontName_S.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_S.ValueChanged += FontSettingsChanged;
            button_Color_S.Click += btnColor_S_Click;

            // Attach event handlers for the Body section.
            comboBox_FontName_B.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_B.ValueChanged += FontSettingsChanged;
            button_Color_B.Click += btnColor_B_Click;

            // Attach event handler for the separate icon color button.
            button_IconColor.Click += btn_IconColor_Click;

            btn_Reset.Click += btn_Reset_Click;
            btn_ReplaceIcons.Click += btn_ReplaceIcons_Click;
            btn_RemoveIcons.Click += btn_RemoveIcons_Click;

            this.Load += (s, e) => panelA4.AutoScrollPosition = new Point(0, 0);
            this.Load += SaveOriginalControlBounds;
            this.Load += A4_Resume_ForceResize;
            this.Load += SetMaximumFormSize;

            // Load icons.
            string iconsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Icons");
            string icon1Path = Path.Combine(iconsFolder, "icon1.png");
            string icon2Path = Path.Combine(iconsFolder, "icon2.png");
            string icon3Path = Path.Combine(iconsFolder, "icon3.png");
            string icon4Path = Path.Combine(iconsFolder, "icon4.png");

            if (File.Exists(icon1Path))
            {
                pictureBox_Icon1.Image = Image.FromFile(icon1Path);
                defaultIcon1 = pictureBox_Icon1.Image;
            }
            if (File.Exists(icon2Path))
            {
                pictureBox_Icon2.Image = Image.FromFile(icon2Path);
                defaultIcon2 = pictureBox_Icon2.Image;
            }
            if (File.Exists(icon3Path))
            {
                pictureBox_Icon3.Image = Image.FromFile(icon3Path);
                defaultIcon3 = pictureBox_Icon3.Image;
            }
            if (File.Exists(icon4Path))
            {
                pictureBox_Icon4.Image = Image.FromFile(icon4Path);
                defaultIcon4 = pictureBox_Icon4.Image;
            }

            // Initialize autosave.
            AttachAutosaveHandlers(this);
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000; // 2 minutes
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();
            this.FormClosing += A4_Resume_FormClosing;

            Task.Delay(100).Wait();
        }

        private void AttachAutosaveHandlers(Control container)
        {
            foreach (Control ctrl in container.Controls)
            {
                if (ctrl is Label || ctrl is TextBox)
                    ctrl.TextChanged += (s, ev) => { unsavedChanges = true; };
                if (ctrl.Controls.Count > 0)
                    AttachAutosaveHandlers(ctrl);
            }
        }

        private void AutosaveTimer_Tick(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();
        }

        private void AutosaveUserSettings()
        {
            // Save label text customizations.
            Properties.Settings.Default.A4_Resume_Name = label_Name.Text;
            Properties.Settings.Default.A4_Resume_PositionTitle = label_PositionTitle.Text;
            Properties.Settings.Default.A4_Resume_Phone = label_Phone.Text;
            Properties.Settings.Default.A4_Resume_Email = label_Email.Text;
            Properties.Settings.Default.A4_Resume_Location = label_Location.Text;
            Properties.Settings.Default.A4_Resume_LinkedIn = label_LinkedIn.Text;
            Properties.Settings.Default.A4_Resume_Head_ProSum = label_ProSummary.Text;
            Properties.Settings.Default.A4_Resume_ProSum = label_ProfessionalSummary.Text;
            Properties.Settings.Default.A4_Resume_Head_Edu = label_Education.Text;
            Properties.Settings.Default.A4_Resume_Deg1 = label_Degree_1.Text;
            Properties.Settings.Default.A4_Resume_Maj1 = label_Major_1.Text;
            Properties.Settings.Default.A4_Resume_Col1 = label_College_1.Text;
            Properties.Settings.Default.A4_Resume_Yr1 = label_Year_1.Text;
            Properties.Settings.Default.A4_Resume_Deg2 = label_Degree_2.Text;
            Properties.Settings.Default.A4_Resume_Maj2 = label_Major_2.Text;
            Properties.Settings.Default.A4_Resume_Col2 = label_College_2.Text;
            Properties.Settings.Default.A4_Resume_Yr2 = label_Year_2.Text;
            Properties.Settings.Default.A4_Resume_Head_Skills = label_Skills.Text;
            Properties.Settings.Default.A4_Resume_SubHead_ProSkills = label_ProSkills.Text;
            Properties.Settings.Default.A4_Resume_PS1 = label_Skill_PS1.Text;
            Properties.Settings.Default.A4_Resume_PS2 = label_Skill_PS2.Text;
            Properties.Settings.Default.A4_Resume_PS3 = label_Skill_PS3.Text;
            Properties.Settings.Default.A4_Resume_PS4 = label_Skill_PS4.Text;
            Properties.Settings.Default.A4_Resume_PS5 = label_Skill_PS5.Text;
            Properties.Settings.Default.A4_Resume_PS6 = label_Skill_PS6.Text;
            Properties.Settings.Default.A4_Resume_PS7 = label_Skill_PS7.Text;
            Properties.Settings.Default.A4_Resume_PS8 = label_Skill_PS8.Text;
            Properties.Settings.Default.A4_Resume_PS9 = label_Skill_PS9.Text;
            Properties.Settings.Default.A4_Resume_SubHead_TechSkills = label_TechSkills.Text;
            Properties.Settings.Default.A4_Resume_TS1 = label_Skill_TS1.Text;
            Properties.Settings.Default.A4_Resume_TS2 = label_Skill_TS2.Text;
            Properties.Settings.Default.A4_Resume_TS3 = label_Skill_TS3.Text;
            Properties.Settings.Default.A4_Resume_TS4 = label_Skill_TS4.Text;
            Properties.Settings.Default.A4_Resume_TS5 = label_Skill_TS5.Text;
            Properties.Settings.Default.A4_Resume_TS6 = label_Skill_TS6.Text;
            Properties.Settings.Default.A4_Resume_TS7 = label_Skill_TS7.Text;
            Properties.Settings.Default.A4_Resume_TS8 = label_Skill_TS8.Text;
            Properties.Settings.Default.A4_Resume_TS9 = label_Skill_TS9.Text;
            Properties.Settings.Default.A4_Resume_Head_WorkExp = label_WorkExp.Text;
            Properties.Settings.Default.A4_Resume_Position1 = label_Position_1.Text;
            Properties.Settings.Default.A4_Resume_Company1 = label_Company_1.Text;
            Properties.Settings.Default.A4_Resume_City1 = label_City_1.Text;
            Properties.Settings.Default.A4_Resume_Year1 = label_Year_1.Text;
            Properties.Settings.Default.A4_Resume_Work1 = label_Work_1.Text;
            Properties.Settings.Default.A4_Resume_Position2 = label_Position_2.Text;
            Properties.Settings.Default.A4_Resume_Company2 = label_Company_2.Text;
            Properties.Settings.Default.A4_Resume_City2 = label_City_2.Text;
            Properties.Settings.Default.A4_Resume_Year2 = label_Year_2.Text;
            Properties.Settings.Default.A4_Resume_Work2 = label_Work_2.Text;

            // Always save customization settings
            Properties.Settings.Default.A4_Resume_FontName_Name = comboBox_FontName_Name.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_Resume_FontSize_Name = (decimal)numericUpDown_Size_Name.Value;
            Properties.Settings.Default.A4_Resume_Color_Name = button_Color_Name.BackColor.ToArgb();

            Properties.Settings.Default.A4_Resume_FontName_PT = comboBox_FontName_PT.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_Resume_FontSize_PT = (decimal)numericUpDown_Size_PT.Value;
            Properties.Settings.Default.A4_Resume_Color_PT = button_Color_PT.BackColor.ToArgb();

            Properties.Settings.Default.A4_Resume_FontName_Heading = comboBox_FontName_H.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_Resume_FontSize_Heading = (decimal)numericUpDown_Size_H.Value;
            Properties.Settings.Default.A4_Resume_Color_Heading = button_Color_H.BackColor.ToArgb();

            Properties.Settings.Default.A4_Resume_FontName_Subheading = comboBox_FontName_S.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_Resume_FontSize_Subheading = (decimal)numericUpDown_Size_S.Value;
            Properties.Settings.Default.A4_Resume_Color_Subheading = button_Color_S.BackColor.ToArgb();

            Properties.Settings.Default.A4_Resume_FontName_Body = comboBox_FontName_B.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_Resume_FontSize_Body = (decimal)numericUpDown_Size_B.Value;
            Properties.Settings.Default.A4_Resume_Color_Body = button_Color_B.BackColor.ToArgb();

            // Save icon file paths.
            Properties.Settings.Default.A4_Resume_Icon1Path = pictureBox_Icon1.Tag as string ?? "";
            Properties.Settings.Default.A4_Resume_Icon2Path = pictureBox_Icon2.Tag as string ?? "";
            Properties.Settings.Default.A4_Resume_Icon3Path = pictureBox_Icon3.Tag as string ?? "";
            Properties.Settings.Default.A4_Resume_Icon4Path = pictureBox_Icon4.Tag as string ?? "";
            // Save the icon color from the separate icon color button.
            Properties.Settings.Default.A4_Resume_IconColor = button_IconColor.BackColor.ToArgb();

            // save icon visibility
            Properties.Settings.Default.A4_Resume_Icon1Visible = pictureBox_Icon1.Visible;
            Properties.Settings.Default.A4_Resume_Icon2Visible = pictureBox_Icon2.Visible;
            Properties.Settings.Default.A4_Resume_Icon3Visible = pictureBox_Icon3.Visible;
            Properties.Settings.Default.A4_Resume_Icon4Visible = pictureBox_Icon4.Visible;

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        private void LoadUserSettings()
        {
            // Load label text customizations.
            label_Name.Text = Properties.Settings.Default.A4_Resume_Name;
            label_PositionTitle.Text = Properties.Settings.Default.A4_Resume_PositionTitle;
            label_Phone.Text = Properties.Settings.Default.A4_Resume_Phone;
            label_Email.Text = Properties.Settings.Default.A4_Resume_Email;
            label_Location.Text = Properties.Settings.Default.A4_Resume_Location;
            label_LinkedIn.Text = Properties.Settings.Default.A4_Resume_LinkedIn;
            label_ProSummary.Text = "PROFESSIONAL SUMMARY";
            // ✅ Headings/Subheadings with safe defaults for Release mode
            label_Education.Text = string.IsNullOrEmpty(Properties.Settings.Default.A4_Resume_Head_Edu)
                ? "EDUCATION"
                : Properties.Settings.Default.A4_Resume_Head_Edu;

            label_Skills.Text = string.IsNullOrEmpty(Properties.Settings.Default.A4_Resume_Head_Skills)
                ? "SKILLS"
                : Properties.Settings.Default.A4_Resume_Head_Skills;

            label_ProSkills.Text = string.IsNullOrEmpty(Properties.Settings.Default.A4_Resume_SubHead_ProSkills)
                ? "PROFESSIONAL SKILLS"
                : Properties.Settings.Default.A4_Resume_SubHead_ProSkills;

            label_TechSkills.Text = string.IsNullOrEmpty(Properties.Settings.Default.A4_Resume_SubHead_TechSkills)
                ? "TECHNICAL SKILLS"
                : Properties.Settings.Default.A4_Resume_SubHead_TechSkills;

            label_WorkExp.Text = string.IsNullOrEmpty(Properties.Settings.Default.A4_Resume_Head_WorkExp)
                ? "WORK EXPERIENCE"
                : Properties.Settings.Default.A4_Resume_Head_WorkExp;

            // Education Section
            label_Degree_1.Text = Properties.Settings.Default.A4_Resume_Deg1;
            label_Major_1.Text = Properties.Settings.Default.A4_Resume_Maj1;
            label_College_1.Text = Properties.Settings.Default.A4_Resume_Col1;
            label_Year_1.Text = Properties.Settings.Default.A4_Resume_Yr1;
            label_Degree_2.Text = Properties.Settings.Default.A4_Resume_Deg2;
            label_Major_2.Text = Properties.Settings.Default.A4_Resume_Maj2;
            label_College_2.Text = Properties.Settings.Default.A4_Resume_Col2;
            label_Year_2.Text = Properties.Settings.Default.A4_Resume_Yr2;

            // Professional Skills
            label_Skill_PS1.Text = Properties.Settings.Default.A4_Resume_PS1;
            label_Skill_PS2.Text = Properties.Settings.Default.A4_Resume_PS2;
            label_Skill_PS3.Text = Properties.Settings.Default.A4_Resume_PS3;
            label_Skill_PS4.Text = Properties.Settings.Default.A4_Resume_PS4;
            label_Skill_PS5.Text = Properties.Settings.Default.A4_Resume_PS5;
            label_Skill_PS6.Text = Properties.Settings.Default.A4_Resume_PS6;
            label_Skill_PS7.Text = Properties.Settings.Default.A4_Resume_PS7;
            label_Skill_PS8.Text = Properties.Settings.Default.A4_Resume_PS8;
            label_Skill_PS9.Text = Properties.Settings.Default.A4_Resume_PS9;

            // Technical Skills
            label_Skill_TS1.Text = Properties.Settings.Default.A4_Resume_TS1;
            label_Skill_TS2.Text = Properties.Settings.Default.A4_Resume_TS2;
            label_Skill_TS3.Text = Properties.Settings.Default.A4_Resume_TS3;
            label_Skill_TS4.Text = Properties.Settings.Default.A4_Resume_TS4;
            label_Skill_TS5.Text = Properties.Settings.Default.A4_Resume_TS5;
            label_Skill_TS6.Text = Properties.Settings.Default.A4_Resume_TS6;
            label_Skill_TS7.Text = Properties.Settings.Default.A4_Resume_TS7;
            label_Skill_TS8.Text = Properties.Settings.Default.A4_Resume_TS8;
            label_Skill_TS9.Text = Properties.Settings.Default.A4_Resume_TS9;

            // Work Experience
            label_Position_1.Text = Properties.Settings.Default.A4_Resume_Position1;
            label_Company_1.Text = Properties.Settings.Default.A4_Resume_Company1;
            label_City_1.Text = Properties.Settings.Default.A4_Resume_City1;
            label_Year_1.Text = Properties.Settings.Default.A4_Resume_Year1;
            label_Work_1.Text = Properties.Settings.Default.A4_Resume_Work1;

            label_Position_2.Text = Properties.Settings.Default.A4_Resume_Position2;
            label_Company_2.Text = Properties.Settings.Default.A4_Resume_Company2;
            label_City_2.Text = Properties.Settings.Default.A4_Resume_City2;
            label_Year_2.Text = Properties.Settings.Default.A4_Resume_Year2;
            label_Work_2.Text = Properties.Settings.Default.A4_Resume_Work2;


            // Name
            string savedFontName = Properties.Settings.Default.A4_Resume_FontName_Name;
            comboBox_FontName_Name.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_Name.Value = Properties.Settings.Default.A4_Resume_FontSize_Name > 0
                ? Properties.Settings.Default.A4_Resume_FontSize_Name : 25;
            int colorVal = Properties.Settings.Default.A4_Resume_Color_Name;
            button_Color_Name.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Position Title
            savedFontName = Properties.Settings.Default.A4_Resume_FontName_PT;
            comboBox_FontName_PT.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_PT.Value = Properties.Settings.Default.A4_Resume_FontSize_PT > 0
                ? Properties.Settings.Default.A4_Resume_FontSize_PT : 12;
            colorVal = Properties.Settings.Default.A4_Resume_Color_PT;
            button_Color_PT.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Headings
            savedFontName = Properties.Settings.Default.A4_Resume_FontName_Heading;
            comboBox_FontName_H.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_H.Value = Properties.Settings.Default.A4_Resume_FontSize_Heading > 0
                ? Properties.Settings.Default.A4_Resume_FontSize_Heading : 12;
            colorVal = Properties.Settings.Default.A4_Resume_Color_Heading;
            button_Color_H.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Subheadings
            savedFontName = Properties.Settings.Default.A4_Resume_FontName_Subheading;
            comboBox_FontName_S.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_S.Value = Properties.Settings.Default.A4_Resume_FontSize_Subheading > 0
                ? Properties.Settings.Default.A4_Resume_FontSize_Subheading : 10;
            colorVal = Properties.Settings.Default.A4_Resume_Color_Subheading;
            button_Color_S.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Body
            savedFontName = Properties.Settings.Default.A4_Resume_FontName_Body;
            comboBox_FontName_B.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_B.Value = Properties.Settings.Default.A4_Resume_FontSize_Body > 0
                ? Properties.Settings.Default.A4_Resume_FontSize_Body : 10;
            colorVal = Properties.Settings.Default.A4_Resume_Color_Body;
            button_Color_B.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Load icon color with fallback
            int iconColorVal = Properties.Settings.Default.A4_Resume_IconColor;
            button_IconColor.BackColor = (iconColorVal == 0) ? Color.Black : Color.FromArgb(iconColorVal);

            // Also update the style so the text disappears like on Reset
            button_IconColor.FlatStyle = FlatStyle.Flat;
            button_IconColor.ForeColor = Color.White;
            button_IconColor.FlatAppearance.BorderColor = Color.White;
            button_IconColor.UseVisualStyleBackColor = false;

            // Apply icon tinting and fonts
            ApplyIconColor(button_IconColor.BackColor);
            ApplyFontChanges();

            string icon1Path = Properties.Settings.Default.A4_Resume_Icon1Path;
            string icon2Path = Properties.Settings.Default.A4_Resume_Icon2Path;
            string icon3Path = Properties.Settings.Default.A4_Resume_Icon3Path;
            string icon4Path = Properties.Settings.Default.A4_Resume_Icon4Path;

            // Load icons with fallback to default if setting is missing or file is missing
            if (!string.IsNullOrEmpty(icon1Path) && File.Exists(icon1Path))
            {
                pictureBox_Icon1.Image = Image.FromFile(icon1Path);
                pictureBox_Icon1.Tag = icon1Path;
            }
            else if (defaultIcon1 != null)
            {
                pictureBox_Icon1.Image = defaultIcon1;
                pictureBox_Icon1.Visible = true;
            }

            if (!string.IsNullOrEmpty(icon2Path) && File.Exists(icon2Path))
            {
                pictureBox_Icon2.Image = Image.FromFile(icon2Path);
                pictureBox_Icon2.Tag = icon2Path;
            }
            else if (defaultIcon2 != null)
            {
                pictureBox_Icon2.Image = defaultIcon2;
                pictureBox_Icon2.Visible = true;
            }

            if (!string.IsNullOrEmpty(icon3Path) && File.Exists(icon3Path))
            {
                pictureBox_Icon3.Image = Image.FromFile(icon3Path);
                pictureBox_Icon3.Tag = icon3Path;
            }
            else if (defaultIcon3 != null)
            {
                pictureBox_Icon3.Image = defaultIcon3;
                pictureBox_Icon3.Visible = true;
            }

            if (!string.IsNullOrEmpty(icon4Path) && File.Exists(icon4Path))
            {
                pictureBox_Icon4.Image = Image.FromFile(icon4Path);
                pictureBox_Icon4.Tag = icon4Path;
            }
            else if (defaultIcon4 != null)
            {
                pictureBox_Icon4.Image = defaultIcon4;
                pictureBox_Icon4.Visible = true;
            }

            // Reapply the saved icon color.
            ApplyIconColor(button_IconColor.BackColor);

            // restore icon visibility
            pictureBox_Icon1.Visible = Properties.Settings.Default.A4_Resume_Icon1Visible;
            pictureBox_Icon2.Visible = Properties.Settings.Default.A4_Resume_Icon2Visible;
            pictureBox_Icon3.Visible = Properties.Settings.Default.A4_Resume_Icon3Visible;
            pictureBox_Icon4.Visible = Properties.Settings.Default.A4_Resume_Icon4Visible;

            label_ProSkills.Visible = true;
            label_Skills.Visible = true;
            label_TechSkills.Visible = true;
            label_WorkExp.Visible = true;
            label_Education.Visible = true;          

        }

        private void A4_Resume_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();
        }

        private IEnumerable<Control> GetAllControls(Control container)
        {
            foreach (Control ctrl in container.Controls)
            {
                yield return ctrl;
                foreach (Control child in GetAllControls(ctrl))
                    yield return child;
            }
        }

        public A4_Resume()
        {
            InitializeComponent();
            string iconsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Icons");
            MessageBox.Show("Icons Folder Path: " + iconsFolder);
            {
                InitializeComponent();
                this.AutoScaleMode = AutoScaleMode.None;
            }
            panelA4.Size = new Size(fixedWidth, fixedHeight);
            this.MinimumSize = new Size(820, 768);
            this.Load += SetMaximumFormSize;
            this.Load += A4_Resume_ForceResize;
        }

        private void SaveOriginalControlBounds(object sender, EventArgs e)
        {
            originalControlBounds.Clear();
            foreach (Control ctrl in panelA4.Controls)
                originalControlBounds[ctrl] = new Rectangle(ctrl.Location, ctrl.Size);
        }

        private void LoadFonts()
        {
            foreach (FontFamily font in FontFamily.Families)
            {
                comboBox_FontName_Name.Items.Add(font.Name);
                comboBox_FontName_PT.Items.Add(font.Name);
                comboBox_FontName_H.Items.Add(font.Name);
                comboBox_FontName_S.Items.Add(font.Name);
                comboBox_FontName_B.Items.Add(font.Name);
            }
            comboBox_FontName_Name.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_PT.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_H.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_S.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_B.SelectedItem = "Microsoft Sans Serif";
        }

        private Font CreateFont(string fontName, float size)
        {
            return new Font(fontName, size);
        }

        private void ApplyFontChanges()
        {
            // For Name section.
            if (comboBox_FontName_Name.SelectedItem != null)
            {
                string fontName = comboBox_FontName_Name.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_Name.Value;
                Color color = button_Color_Name.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Name"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }
            // For Position Title section.
            if (comboBox_FontName_PT.SelectedItem != null)
            {
                string fontName = comboBox_FontName_PT.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_PT.Value;
                Color color = button_Color_PT.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "PositionTitle"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }
            // For Headings.
            if (comboBox_FontName_H.SelectedItem != null)
            {
                string fontName = comboBox_FontName_H.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_H.Value;
                Color color = button_Color_H.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Heading"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }
            // For Subheadings.
            if (comboBox_FontName_S.SelectedItem != null)
            {
                string fontName = comboBox_FontName_S.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_S.Value;
                Color color = button_Color_S.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Subheading"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }
            // For Body.
            if (comboBox_FontName_B.SelectedItem != null)
            {
                string fontName = comboBox_FontName_B.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_B.Value;
                Color color = button_Color_B.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }
        }

        private void FontSettingsChanged(object sender, EventArgs e)
        {
            ApplyFontChanges();
            unsavedChanges = true;
        }

        private void btnFontColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_Color_Name.BackColor = colorDialog.Color;
                    ApplyFontChanges();
                    unsavedChanges = true;
                }
            }
        }

        private void btnColor_PT_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_Color_PT.BackColor = colorDialog.Color;
                    ApplyFontChanges();
                    unsavedChanges = true;
                }
            }
        }

        private void btnColor_H_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_Color_H.BackColor = colorDialog.Color;
                    ApplyFontChanges();
                    unsavedChanges = true;
                }
            }
        }

        private void btnColor_S_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_Color_S.BackColor = colorDialog.Color;
                    ApplyFontChanges();
                    unsavedChanges = true;
                }
            }
        }

        private void btnColor_B_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_Color_B.BackColor = colorDialog.Color;
                    ApplyFontChanges();
                    unsavedChanges = true;
                }
            }
        }

        private void btn_IconColor_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button_IconColor.BackColor = colorDialog.Color;
                    ApplyIconColor(colorDialog.Color);
                    unsavedChanges = true;
                    AutosaveUserSettings();  // Force saving so the new icon color is saved immediately
                }
            }
        }

        private void ApplyIconColor(Color newColor)
        {
            PictureBox[] icons = { pictureBox_Icon1, pictureBox_Icon2, pictureBox_Icon3, pictureBox_Icon4 };
            foreach (PictureBox icon in icons)
            {
                if (icon.Image != null)
                {
                    Bitmap coloredIcon = new Bitmap(icon.Image);
                    for (int y = 0; y < coloredIcon.Height; y++)
                    {
                        for (int x = 0; x < coloredIcon.Width; x++)
                        {
                            Color pixelColor = coloredIcon.GetPixel(x, y);
                            if (pixelColor.A > 0)
                            {
                                coloredIcon.SetPixel(x, y, Color.FromArgb(pixelColor.A, newColor.R, newColor.G, newColor.B));
                            }
                        }
                    }
                    icon.Image = coloredIcon;
                }
            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            // Update UI controls to desired default values:
            comboBox_FontName_Name.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_Name.Value = 25; // Name: 25
            button_Color_Name.BackColor = Color.Black;

            comboBox_FontName_PT.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_PT.Value = 12; // Position Title: 12
            button_Color_PT.BackColor = Color.Black;

            comboBox_FontName_H.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_H.Value = 12; // Headings: 12
            button_Color_H.BackColor = Color.Black;

            comboBox_FontName_S.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_S.Value = 10; // Subheadings: 10
            button_Color_S.BackColor = Color.Black;

            comboBox_FontName_B.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_B.Value = 10; // Body: 10
            button_Color_B.BackColor = Color.Black;

            // Now update the resume template labels by iterating over controls by Tag.
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Name"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 25, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "PositionTitle"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Heading"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Subheading"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
        }

        // Reset Icons Button Functionality
        private void btn_ResetIcons_Click(object sender, EventArgs e)
        {
            // Reset the icon color button to black.
            button_IconColor.UseVisualStyleBackColor = false;  // ensure custom colors are used
            button_IconColor.FlatStyle = FlatStyle.Flat;         // set flat style to show the BackColor
            button_IconColor.BackColor = Color.Black;
            button_IconColor.ForeColor = Color.White;            // adjust text color for contrast (optional)

            // Reset each icon to its default image, ensure it is visible, and clear its Tag.
            if (defaultIcon1 != null)
            {
                pictureBox_Icon1.Image = defaultIcon1;
                pictureBox_Icon1.Visible = true;
                pictureBox_Icon1.Tag = "";
            }
            if (defaultIcon2 != null)
            {
                pictureBox_Icon2.Image = defaultIcon2;
                pictureBox_Icon2.Visible = true;
                pictureBox_Icon2.Tag = "";
            }
            if (defaultIcon3 != null)
            {
                pictureBox_Icon3.Image = defaultIcon3;
                pictureBox_Icon3.Visible = true;
                pictureBox_Icon3.Tag = "";
            }
            if (defaultIcon4 != null)
            {
                pictureBox_Icon4.Image = defaultIcon4;
                pictureBox_Icon4.Visible = true;
                pictureBox_Icon4.Tag = "";
            }
            unsavedChanges = true;
        }

        private void btn_ReplaceIcons_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox_Icon1.Image = Image.FromFile(openFileDialog.FileName);
                    ApplyIconColor(button_IconColor.BackColor); // ✅ Apply current color to new icon
                    pictureBox_Icon1.Tag = openFileDialog.FileName;
                    unsavedChanges = true;
                }
            }
        }

        private void btn_RemoveIcons_Click(object sender, EventArgs e)
        {
            // Get all PictureBox controls whose names start with "pictureBox_Icon"
            var icons = GetAllControls(panelA4).OfType<PictureBox>()
                        .Where(x => x.Name.StartsWith("pictureBox_Icon"))
                        .ToList();
            MessageBox.Show("Found " + icons.Count + " icons.");
            foreach (PictureBox pic in icons)
                pic.Visible = false;
            unsavedChanges = true;
        }

        private void PrintPreviewDialog()
        {
            using (PrintPreviewDialog previewDialog = new PrintPreviewDialog())
            {
                previewDialog.Document = printDocument;
                previewDialog.WindowState = FormWindowState.Maximized;
                previewDialog.ShowDialog();
            }
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            int pageWidth = printDocument.DefaultPageSettings.PaperSize.Width;
            int pageHeight = printDocument.DefaultPageSettings.PaperSize.Height;

            // White background
            e.Graphics.FillRectangle(Brushes.White, 0, 0, pageWidth, pageHeight);

            // Reset any scroll
            if (panel1.AutoScroll)
                panel1.AutoScrollPosition = new Point(0, 0);
            panel1.Refresh();

            // Temporarily force white background on panel
            Color originalBackColor = panel1.BackColor;
            panel1.BackColor = Color.White;

            // 1) Record and hide all Labels and PictureBoxes
            var tempStates = new List<Tuple<Control, bool>>();
            foreach (Control ctrl in panel1.Controls)
            {
                if (ctrl is Label || ctrl is PictureBox)
                {
                    tempStates.Add(Tuple.Create(ctrl, ctrl.Visible));
                    ctrl.Visible = false;
                }
            }

            // Capture the “clean” panel bitmap
            var bmp = new Bitmap(panel1.Width, panel1.Height);
            bmp.SetResolution(e.Graphics.DpiX, e.Graphics.DpiY);
            panel1.DrawToBitmap(bmp, new Rectangle(0, 0, panel1.Width, panel1.Height));

            // 2) Restore each control’s original Visible state
            foreach (var state in tempStates)
                state.Item1.Visible = state.Item2;

            // Restore panel background
            panel1.BackColor = originalBackColor;

            // Compute margins for centering
            int leftMargin = (pageWidth - panel1.Width) / 2;
            int topMargin = (pageHeight - panel1.Height) / 2;

            // Draw the captured background
            e.Graphics.DrawImage(bmp, leftMargin, topMargin, panel1.Width, panel1.Height);

            // High‑quality rendering
            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Draw all labels as vector text
            foreach (Control ctrl in panel1.Controls)
            {
                if (ctrl is Label lbl)
                {
                    var printPosX = leftMargin + lbl.Left;
                    var printPosY = topMargin + lbl.Top;
                    var textRect = new RectangleF(printPosX, printPosY, lbl.Width, lbl.Height);

                    using (var brush = new SolidBrush(lbl.ForeColor))
                        e.Graphics.DrawString(lbl.Text, lbl.Font, brush, textRect,
                            new StringFormat { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Near });
                }
            }

            // 3) Draw only the icons that remain Visible
            if (pictureBox_Icon1.Visible)
                DrawIconNextToLabel(e.Graphics, pictureBox_Icon1.Image, label_Phone, -22, 0, leftMargin, topMargin);
            if (pictureBox_Icon2.Visible)
                DrawIconNextToLabel(e.Graphics, pictureBox_Icon2.Image, label_Email, -22, 1, leftMargin, topMargin);
            if (pictureBox_Icon3.Visible)
                DrawIconNextToLabel(e.Graphics, pictureBox_Icon3.Image, label_Location, -22, -1, leftMargin, topMargin);
            if (pictureBox_Icon4.Visible)
                DrawIconNextToLabel(e.Graphics, pictureBox_Icon4.Image, label_LinkedIn, -22, -2, leftMargin, topMargin);

            bmp.Dispose();
            e.HasMorePages = false;
        }


        // Helper method for precise alignment
        private void DrawIconNextToLabel(Graphics g, Image icon, Label lbl, int offsetX, int offsetY, int leftMargin, int topMargin)
        {
            if (icon == null || lbl == null) return;

            // Center vertically with slight manual adjustment (offsetY)
            float iconX = leftMargin + lbl.Left + offsetX;
            float iconY = topMargin + lbl.Top + offsetY;

            // Draw icons slightly smaller (16x16 instead of 18x18) for sharpness:
            g.DrawImage(icon, iconX, iconY, 16, 16);
        }

        private void PrintTextVector(Graphics g, int leftMargin, int topMargin)
        {
            foreach (Control ctrl in panel1.Controls)
            {
                if (ctrl is Label lbl)
                {
                    // Calculate the absolute position of the label relative to panel1.
                    // (Assuming no nested container offset—if there is, you may need to adjust.)
                    float printPosX = leftMargin + lbl.Left;
                    float printPosY = topMargin + lbl.Top;

                    // Draw the label's text using its font and color.
                    using (SolidBrush brush = new SolidBrush(lbl.ForeColor))
                    {
                        g.DrawString(lbl.Text, lbl.Font, brush, new PointF(printPosX, printPosY));
                    }
                }
            }
        }

        private void A4_Resume_Load(object sender, EventArgs e)
        {
            panelA4.Location = new Point((this.ClientSize.Width - panelA4.Width) / 2, panelControls.Bottom + 10);

            // ✅ Force label visibility
            label_ProSkills.Visible = true;
            label_Skills.Visible = true;
            label_TechSkills.Visible = true;
            label_WorkExp.Visible = true;
            label_Education.Visible = true;

            // ✅ Set initial color values for all font and icon color buttons
            button_Color_Name.BackColor = Color.Black;
            button_Color_PT.BackColor = Color.Black;
            button_Color_H.BackColor = Color.Black;
            button_Color_S.BackColor = Color.Black;
            button_Color_B.BackColor = Color.Black;
            button_IconColor.BackColor = Color.Black;
        }

        private void A4_Resume_Resize(object sender, EventArgs e)
        {
            panelA4.Size = new Size(fixedWidth, fixedHeight);
            panelA4.Location = new Point((this.ClientSize.Width - panelA4.Width) / 2, panelControls.Bottom + 10);
        }

        private void A4_Resume_ForceResize(object sender, EventArgs e)
        {
            A4_Resume_Resize(this, EventArgs.Empty);
        }

        private void SetMaximumFormSize(object sender, EventArgs e)
        {
            if (panelControls != null && panelA4 != null)
            {
                int maxHeight = panelControls.Height + fixedHeight + SystemInformation.CaptionHeight + (SystemInformation.BorderSize.Height * 2);
                this.MaximumSize = new Size(820, maxHeight);
            }
        }

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

        private void button_Back_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var workExp = new A4_Form_WorkExperience();
            workExp.Size = this.Size;
            workExp.Show();
            this.Close();
        }

        private void button_Next_CL_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var coverLetterSize = new CoverLetter_Size();
            coverLetterSize.Size = this.Size;
            coverLetterSize.Show();
            this.Close();
        }

        private void btn_ReplaceIcon_2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox_Icon2.Image = Image.FromFile(openFileDialog.FileName);
                    ApplyIconColor(button_IconColor.BackColor);
                    pictureBox_Icon2.Tag = openFileDialog.FileName;                    
                    unsavedChanges = true;
                }
            }
        }

        private void btn_ReplaceIcon_3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox_Icon3.Image = Image.FromFile(openFileDialog.FileName);
                    ApplyIconColor(button_IconColor.BackColor);
                    pictureBox_Icon3.Tag = openFileDialog.FileName;
                    unsavedChanges = true;
                }
            }
        }

        private void btn_ReplaceIcon_4_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox_Icon4.Image = Image.FromFile(openFileDialog.FileName);
                    ApplyIconColor(button_IconColor.BackColor);
                    pictureBox_Icon4.Tag = openFileDialog.FileName;
                    unsavedChanges = true;
                }
            }
        }

        private void button_Save_Click(object sender, EventArgs e)
        {
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pictureBox_Print_Click(object sender, EventArgs e)
        {
            // Create a PrintDialog and assign the PrintDocument.
            PrintDialog printDlg = new PrintDialog();
            printDlg.Document = printDocument;  // This document now has A4 as its default paper size.

            // Show the dialog; if the user clicks OK, proceed with printing.
            if (printDlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    printDocument.Print();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while printing: " + ex.Message,
                                    "Printing Error",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error);
                }
            }
        }
    }
}

using AutoFill_Resume_Pro.Forms;
using AutoFill_Resume_Pro.Selection_Pages;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;

namespace AutoFill_Resume_Pro.Templates
{
    public partial class US_CoverLetter : Form
    {
        // Constants for fixed US Letter dimensions.
        private const int fixedWidth = 816;
        private const int fixedHeight = 1056;

        private PrintDocument printDocument;
        private Dictionary<Control, Rectangle> originalControlBounds = new Dictionary<Control, Rectangle>();

        private Image defaultIcon1;
        private Image defaultIcon2;
        private Image defaultIcon3;
        private Image defaultIcon4;

        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        public US_CoverLetter(string[] userData)
        {
            // 1) Base sizing & scrolling
            this.AutoScaleMode = AutoScaleMode.None;
            this.ClientSize = new Size(832, 1350);
            this.MinimumSize = new Size(832, 768);
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 950);

            InitializeComponent();
            
            panelUS.BackColor = Color.White;

            // Fixed US‑Letter panel and rendering
            panelUS.Size = new Size(fixedWidth, fixedHeight);
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.Resize += US_CoverLetter_Resize;

            // 2) Centralize all mapping in one Load handler
            this.Load += (s, e) =>
            {
                // a) Apply saved customizations
                LoadUserSettings();

                // b) Override with incoming userData (needs ≥ 14 items)
                if (userData != null && userData.Length >= 14)
                {
                    // Header (0–5)
                    label_Name.Text = userData[0];
                    label_PositionTitle.Text = userData[1];
                    label_Phone.Text = userData[2];
                    label_Email.Text = userData[3];
                    label_Location.Text = userData[4];
                    label_LinkedIn.Text = userData[5];

                    // Cover‑letter fields (6–13)
                    label_Date.Text = userData[6];
                    label_HMgrsFName.Text = userData[7];
                    label_CompanyName.Text = userData[8];
                    label_CompanyAddress.Text = userData[9];
                    label_City.Text = userData[10];
                    label_HMgrsN.Text = userData[11];
                    label_CL_Body.Text = userData[12];
                    label_YourName.Text = userData[13];
                }
            };           

        printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);

            PaperSize usSize = new PaperSize("US Letter", 850, 1100); // 8.5in x 11in * 100
            printDocument.DefaultPageSettings.PaperSize = usSize;
            printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            this.Resize += US_CoverLetter_Resize;

            // Load fonts into all ComboBoxes.
            LoadFonts();

            // Set default font size and color for all sections.
            numericUpDown_Size_Name.Value = 25; // Name
            numericUpDown_Size_PT.Value = 12;     // Position Title
            numericUpDown_Size_B.Value = 10;      // Body
            numericUpDown_Size_S.Value = 20;      // Signature

            button_Color_Name.BackColor = Color.Black;
            button_Color_PT.BackColor = Color.Black;
            button_Color_B.BackColor = Color.Black;
            button_Color_S.BackColor = Color.Black;
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

            // Attach event handlers for the Body section.
            comboBox_FontName_B.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_B.ValueChanged += FontSettingsChanged;
            button_Color_B.Click += btnColor_B_Click;

            // Attach event handlers for the Signature section.
            comboBox_FontName_S.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_S.ValueChanged += FontSettingsChanged;
            button_Color_S.Click += btnColor_S_Click;

            // Attach event handler for the separate icon color button.
            button_IconColor.Click += btn_IconColor_Click;

            // Attach event handlers for reset, replace, and remove icons.
            btn_Reset.Click += btn_Reset_Click;
            btn_RemoveIcons.Click += btn_RemoveIcons_Click;

            // Additional load event handlers.
            this.Load += (s, e) => panelUS.AutoScrollPosition = new Point(0, 0);
            this.Load += SaveOriginalControlBounds;
            this.Load += SetMaximumFormSize;

            // Load icons from the Resources/Icons folder.
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
            this.FormClosing += US_CoverLetter_FormClosing;
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var usInfo = new US_CoverLetter_Info();
            usInfo.Size = this.Size;
            usInfo.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();

            var rls = new ReferenceLetter_Size();
            rls.Size = this.Size;
            rls.Show();
            this.Close();
        }

        private void US_CoverLetter_Resize(object sender, EventArgs e)
        {
            panelUS.Size = new Size(fixedWidth, fixedHeight);
            if (panelControls != null)
            {
                panelUS.Location = new Point((this.ClientSize.Width - panelUS.Width) / 2, panelControls.Bottom + 10);
            }
        }

        private void SetMaximumFormSize(object sender, EventArgs e)
        {
            if (panelControls != null && panelUS != null)
            {
                int maxHeight = panelControls.Height + fixedHeight + SystemInformation.CaptionHeight + (SystemInformation.BorderSize.Height * 2);
                this.MaximumSize = new Size(840, maxHeight);
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

        private void DrawIconNextToLabel(Graphics g, Image icon, Label lbl, int offsetX, int offsetY, int leftMargin, int topMargin)
        {
            if (icon == null || lbl == null)
                return;

            float iconX = leftMargin + lbl.Left + offsetX;
            float iconY = topMargin + lbl.Top + offsetY;

            g.DrawImage(icon, iconX, iconY, 16, 16);
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
            Properties.Settings.Default.US_CoverLetter_Name = label_Name.Text;
            Properties.Settings.Default.US_CoverLetter_PositionTitle = label_PositionTitle.Text;
            Properties.Settings.Default.US_CoverLetter_Phone = label_Phone.Text;
            Properties.Settings.Default.US_CoverLetter_Email = label_Email.Text;
            Properties.Settings.Default.US_CoverLetter_Location = label_Location.Text;
            Properties.Settings.Default.US_CoverLetter_LinkedIn = label_LinkedIn.Text;

            Properties.Settings.Default.US_CoverLetter_Date = label_Date.Text;
            Properties.Settings.Default.US_CoverLetter_HiringMgr = label_HMgrsFName.Text;
            Properties.Settings.Default.US_CoverLetter_Company = label_CompanyName.Text;
            Properties.Settings.Default.US_CoverLetter_CompAddress = label_CompanyAddress.Text;
            Properties.Settings.Default.US_CoverLetter_City = label_City.Text;
            Properties.Settings.Default.US_CoverLetter_HMN = label_HMgrsN.Text;
            Properties.Settings.Default.US_CoverLetter_Body = label_CL_Body.Text;
            Properties.Settings.Default.US_CoverLetter_YFN = label_YourName.Text;

            Properties.Settings.Default.US_CoverLetter_FontName_Name = comboBox_FontName_Name.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.US_CoverLetter_FontSize_Name = (decimal)numericUpDown_Size_Name.Value;
            Properties.Settings.Default.US_CoverLetter_Color_Name = button_Color_Name.BackColor.ToArgb();

            Properties.Settings.Default.US_CoverLetter_FontName_PT = comboBox_FontName_PT.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.US_CoverLetter_FontSize_PT = (decimal)numericUpDown_Size_PT.Value;
            Properties.Settings.Default.US_CoverLetter_Color_PT = button_Color_PT.BackColor.ToArgb();

            Properties.Settings.Default.US_CoverLetter_FontName_Body = comboBox_FontName_B.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.US_CoverLetter_FontSize_Body = (decimal)numericUpDown_Size_B.Value;
            Properties.Settings.Default.US_CoverLetter_Color_Body = button_Color_B.BackColor.ToArgb();

            Properties.Settings.Default.US_CoverLetter_FontName_Signature = comboBox_FontName_S.SelectedItem?.ToString() ?? "Monotype Corsiva";
            Properties.Settings.Default.US_CoverLetter_FontSize_Signature = (decimal)numericUpDown_Size_S.Value;
            Properties.Settings.Default.US_CoverLetter_Color_Signature = button_Color_S.BackColor.ToArgb();

            Properties.Settings.Default.US_CoverLetter_Icon1Path = pictureBox_Icon1.Tag as string ?? "";
            Properties.Settings.Default.US_CoverLetter_Icon2Path = pictureBox_Icon2.Tag as string ?? "";
            Properties.Settings.Default.US_CoverLetter_Icon3Path = pictureBox_Icon3.Tag as string ?? "";
            Properties.Settings.Default.US_CoverLetter_Icon4Path = pictureBox_Icon4.Tag as string ?? "";
            Properties.Settings.Default.US_CoverLetter_IconColor = button_IconColor.BackColor.ToArgb();

            // Persist each icon’s Visible state:
            Properties.Settings.Default.US_CoverLetter_Icon1Visible = pictureBox_Icon1.Visible;
            Properties.Settings.Default.US_CoverLetter_Icon2Visible = pictureBox_Icon2.Visible;
            Properties.Settings.Default.US_CoverLetter_Icon3Visible = pictureBox_Icon3.Visible;
            Properties.Settings.Default.US_CoverLetter_Icon4Visible = pictureBox_Icon4.Visible;

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        private void LoadUserSettings()
        {
            label_Name.Text = Properties.Settings.Default.US_CoverLetter_Name;
            label_PositionTitle.Text = Properties.Settings.Default.US_CoverLetter_PositionTitle;
            label_Phone.Text = Properties.Settings.Default.US_CoverLetter_Phone;
            label_Email.Text = Properties.Settings.Default.US_CoverLetter_Email;
            label_Location.Text = Properties.Settings.Default.US_CoverLetter_Location;
            label_LinkedIn.Text = Properties.Settings.Default.US_CoverLetter_LinkedIn;
            label_Date.Text = Properties.Settings.Default.US_CoverLetter_Date;
            label_HMgrsFName.Text = Properties.Settings.Default.US_CoverLetter_HiringMgr;
            label_CompanyName.Text = Properties.Settings.Default.US_CoverLetter_Company;
            label_CompanyAddress.Text = Properties.Settings.Default.US_CoverLetter_CompAddress;
            label_City.Text = Properties.Settings.Default.US_CoverLetter_City;
            label_HMgrsN.Text = Properties.Settings.Default.US_CoverLetter_HMN;
            label_CL_Body.Text = Properties.Settings.Default.US_CoverLetter_Body;
            label_YourName.Text = Properties.Settings.Default.US_CoverLetter_YFN;

            // Name
            string savedFontName = Properties.Settings.Default.US_CoverLetter_FontName_Name;
            comboBox_FontName_Name.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_Name.Value = Properties.Settings.Default.US_CoverLetter_FontSize_Name > 0
                ? Properties.Settings.Default.US_CoverLetter_FontSize_Name : 25;
            int colorVal = Properties.Settings.Default.US_CoverLetter_Color_Name;
            button_Color_Name.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Position Title
            savedFontName = Properties.Settings.Default.US_CoverLetter_FontName_PT;
            comboBox_FontName_PT.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_PT.Value = Properties.Settings.Default.US_CoverLetter_FontSize_PT > 0
                ? Properties.Settings.Default.US_CoverLetter_FontSize_PT : 12;
            colorVal = Properties.Settings.Default.US_CoverLetter_Color_PT;
            button_Color_PT.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Body
            savedFontName = Properties.Settings.Default.US_CoverLetter_FontName_Body;
            comboBox_FontName_B.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_B.Value = Properties.Settings.Default.US_CoverLetter_FontSize_Body > 0
                ? Properties.Settings.Default.US_CoverLetter_FontSize_Body : 10;
            colorVal = Properties.Settings.Default.US_CoverLetter_Color_Body;
            button_Color_B.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Signature
            savedFontName = Properties.Settings.Default.US_CoverLetter_FontName_Signature;
            comboBox_FontName_S.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Monotype Corsiva";
            numericUpDown_Size_S.Value = Properties.Settings.Default.US_CoverLetter_FontSize_Signature > 0
                ? Properties.Settings.Default.US_CoverLetter_FontSize_Signature : 20;
            colorVal = Properties.Settings.Default.US_CoverLetter_Color_Signature;
            button_Color_S.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Icon color (with fallback)
            int iconColorVal = Properties.Settings.Default.US_CoverLetter_IconColor;
            button_IconColor.BackColor = (iconColorVal == 0) ? Color.Black : Color.FromArgb(iconColorVal);

            // Match reset styling
            button_IconColor.FlatStyle = FlatStyle.Flat;
            button_IconColor.ForeColor = Color.White;
            button_IconColor.FlatAppearance.BorderColor = Color.White;
            button_IconColor.UseVisualStyleBackColor = false;

            // Apply settings
            ApplyIconColor(button_IconColor.BackColor);
            ApplyFontChanges();

            string icon1Path = Properties.Settings.Default.US_CoverLetter_Icon1Path;
            string icon2Path = Properties.Settings.Default.US_CoverLetter_Icon2Path;
            string icon3Path = Properties.Settings.Default.US_CoverLetter_Icon3Path;
            string icon4Path = Properties.Settings.Default.US_CoverLetter_Icon4Path;

            if (!string.IsNullOrEmpty(icon1Path) && File.Exists(icon1Path))
            {
                pictureBox_Icon1.Image = Image.FromFile(icon1Path);
                pictureBox_Icon1.Tag = icon1Path;
                ApplyIconColor(button_IconColor.BackColor);
            }
            if (!string.IsNullOrEmpty(icon2Path) && File.Exists(icon2Path))
            {
                pictureBox_Icon2.Image = Image.FromFile(icon2Path);
                pictureBox_Icon2.Tag = icon2Path;
                ApplyIconColor(button_IconColor.BackColor);
            }
            if (!string.IsNullOrEmpty(icon3Path) && File.Exists(icon3Path))
            {
                pictureBox_Icon3.Image = Image.FromFile(icon3Path);
                pictureBox_Icon3.Tag = icon3Path;
                ApplyIconColor(button_IconColor.BackColor);
            }
            if (!string.IsNullOrEmpty(icon4Path) && File.Exists(icon4Path))
            {
                pictureBox_Icon4.Image = Image.FromFile(icon4Path);
                pictureBox_Icon4.Tag = icon4Path;
                ApplyIconColor(button_IconColor.BackColor);
            }

            // Restore icon visibility:
            pictureBox_Icon1.Visible = Properties.Settings.Default.US_CoverLetter_Icon1Visible;
            pictureBox_Icon2.Visible = Properties.Settings.Default.US_CoverLetter_Icon2Visible;
            pictureBox_Icon3.Visible = Properties.Settings.Default.US_CoverLetter_Icon3Visible;
            pictureBox_Icon4.Visible = Properties.Settings.Default.US_CoverLetter_Icon4Visible;
        }

        // Helper method to get all child controls.
        private IEnumerable<Control> GetAllControls(Control container)
        {
            foreach (Control ctrl in container.Controls)
            {
                yield return ctrl;
                foreach (Control child in GetAllControls(ctrl))
                    yield return child;
            }
        }

        // Save the original positions and sizes of controls inside panelA4.
        private void SaveOriginalControlBounds(object sender, EventArgs e)
        {
            originalControlBounds.Clear();
            foreach (Control ctrl in panelUS.Controls)
                originalControlBounds[ctrl] = new Rectangle(ctrl.Location, ctrl.Size);
        }

        // FONT CUSTOMIZATION METHODS

        private void LoadFonts()
        {
            foreach (FontFamily font in FontFamily.Families)
            {
                comboBox_FontName_Name.Items.Add(font.Name);
                comboBox_FontName_PT.Items.Add(font.Name);
                comboBox_FontName_B.Items.Add(font.Name);
                comboBox_FontName_S.Items.Add(font.Name);
            }
            comboBox_FontName_Name.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_PT.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_B.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_S.SelectedItem = "Monotype Corsiva";
        }

        private Font CreateFont(string fontName, float size)
        {
            return new Font(fontName, size);
        }

        private void ApplyFontChanges()
        {
            if (comboBox_FontName_Name.SelectedItem != null)
            {
                string fontName = comboBox_FontName_Name.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_Name.Value;
                Color color = button_Color_Name.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Name"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }

            if (comboBox_FontName_PT.SelectedItem != null)
            {
                string fontName = comboBox_FontName_PT.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_PT.Value;
                Color color = button_Color_PT.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "PositionTitle"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }

            if (comboBox_FontName_B.SelectedItem != null)
            {
                string fontName = comboBox_FontName_B.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_B.Value;
                Color color = button_Color_B.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }

            if (comboBox_FontName_S.SelectedItem != null)
            {
                string fontName = comboBox_FontName_S.SelectedItem.ToString();
                float size = (float)numericUpDown_Size_S.Value;
                Color color = button_Color_S.BackColor;
                Font newFont = CreateFont(fontName, size);
                foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Signature"))
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
                    AutosaveUserSettings();
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
                            Color pixel = coloredIcon.GetPixel(x, y);
                            if (pixel.A > 0)
                                coloredIcon.SetPixel(x, y, Color.FromArgb(pixel.A, newColor.R, newColor.G, newColor.B));
                        }
                    }
                    icon.Image = coloredIcon;
                }
            }
        }

        private void btn_Reset_Click(object sender, EventArgs e)
        {
            // Reset font settings to defaults.
            comboBox_FontName_Name.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_Name.Value = 25;
            button_Color_Name.BackColor = Color.Black;

            comboBox_FontName_PT.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_PT.Value = 12;
            button_Color_PT.BackColor = Color.Black;

            comboBox_FontName_B.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_B.Value = 10;
            button_Color_B.BackColor = Color.Black;

            comboBox_FontName_S.SelectedItem = "Monotype Corsiva";
            numericUpDown_Size_S.Value = 20;
            button_Color_S.BackColor = Color.Black;

            foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Name"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 25, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "PositionTitle"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Signature"))
            {
                lbl.Font = new Font("Monotype Corsiva", 20, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
            foreach (Label lbl in GetAllControls(panelUS).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
        }

        private void btn_ResetIcons_Click_1(object sender, EventArgs e)
        {
            // Reset the icon color button to black.
            button_IconColor.UseVisualStyleBackColor = false;
            button_IconColor.FlatStyle = FlatStyle.Flat;
            button_IconColor.BackColor = Color.Black;
            button_IconColor.ForeColor = Color.White;

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

        private void US_CoverLetter_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();
        }

        private void btn_RemoveIcons_Click(object sender, EventArgs e)
        {
            var icons = GetAllControls(panelUS).OfType<PictureBox>()
                        .Where(x => x.Name.StartsWith("pictureBox_Icon"))
                        .ToList();
            MessageBox.Show("Found " + icons.Count + " icons.");
            foreach (PictureBox pic in icons)
                pic.Visible = false;
            unsavedChanges = true;
        }

        private void btn_ReplaceIcons_Click_1(object sender, EventArgs e)
        {        
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    pictureBox_Icon1.Image = Image.FromFile(openFileDialog.FileName);
                    ApplyIconColor(button_IconColor.BackColor);
                    pictureBox_Icon1.Tag = openFileDialog.FileName;
                    unsavedChanges = true;
                }
            }
        }

        // Additional Replace Icon event handlers.
        private void btn_ReplaceIcon_2_Click_1(object sender, EventArgs e)
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

        private void btn_ReplaceIcon_3_Click_1(object sender, EventArgs e)
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

        private void btn_ReplaceIcon_4_Click_1(object sender, EventArgs e)
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

        private void button_Save_Click_1(object sender, EventArgs e)
        {        
            AutosaveUserSettings();
            MessageBox.Show("Saved successfully!", "Save", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // (Optional) PrintPreviewDialog method (only one definition).
        private void PrintPreviewDialog()
        {
            using (PrintPreviewDialog previewDialog = new PrintPreviewDialog())
            {
                previewDialog.Document = printDocument;
                previewDialog.WindowState = FormWindowState.Maximized;
                previewDialog.ShowDialog();
            }
        }

        // (Optional) Helper method to draw text as vector graphics.
        private void PrintTextVector(Graphics g, int leftMargin, int topMargin)
        {
            foreach (Control ctrl in panel1.Controls)
            {
                if (ctrl is Label lbl)
                {
                    float printPosX = leftMargin + lbl.Left;
                    float printPosY = topMargin + lbl.Top;
                    using (SolidBrush brush = new SolidBrush(lbl.ForeColor))
                    {
                        g.DrawString(lbl.Text, lbl.Font, brush, new PointF(printPosX, printPosY));
                    }
                }
            }
        }

        private void US_CoverLetter_Load(object sender, EventArgs e)
        {
            if (panelControls != null)
            {
                panelUS.Location = new Point((this.ClientSize.Width - panelUS.Width) / 2, panelControls.Bottom + 10);
            }
        }

        private void pictureBox_Print_Click(object sender, EventArgs e)
        {
            // Create a PrintDialog and assign the PrintDocument.
            PrintDialog printDlg = new PrintDialog();
            printDlg.Document = printDocument;

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

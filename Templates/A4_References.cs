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
    public partial class A4_References : Form
    {
        // Constants for fixed A4 dimensions.
        private const int fixedWidth = 794;
        private const int fixedHeight = 1123;

        private PrintDocument printDocument;
        // Store the original bounds of controls in panelA4.
        private Dictionary<Control, Rectangle> originalControlBounds = new Dictionary<Control, Rectangle>();

        // Default icon images.
        private Image defaultIcon1;
        private Image defaultIcon2;
        private Image defaultIcon3;
        private Image defaultIcon4;

        // Autosave-related fields.
        private Timer autosaveTimer;
        private bool unsavedChanges = false;

        // Constructor that accepts userData.
        public A4_References(string[] userData)
        {
            // Disable auto scaling to prevent unexpected DPI changes.
            this.AutoScaleMode = AutoScaleMode.None;

            // Set the exact print size (A4: 794x1123 pixels at 96 DPI)
            this.ClientSize = new Size(810, 1350);
            this.MinimumSize = new Size(810, 768);
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 950);

            InitializeComponent(); // Load designer elements

            panelA4.BackColor = Color.White;

            panelA4.Size = new Size(fixedWidth, fixedHeight);
            this.DoubleBuffered = true;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;

            this.Resize += A4_References_Resize;
            this.Load += A4_References_Resize;

            // Apply userData AFTER controls are fully loaded
            this.Load += (s, e) =>
            {
                LoadUserSettings(); // Load fonts/colors first

                if (userData != null && userData.Length >= 31)
                {
                    // Always override with Form1 data
                    label_Name.Text = userData[0];
                    label_PositionTitle.Text = userData[1];
                    label_Phone.Text = userData[2];
                    label_Email.Text = userData[3];
                    label_Location.Text = userData[4];
                    label_LinkedIn.Text = userData[5];

                    // Reference entries
                    lbl_Ref1_N.Text = userData[6];
                    lbl_Ref1_PT.Text = userData[7];
                    lbl_Ref1_Org.Text = userData[8];
                    lbl_Ref1_Ph.Text = userData[9];
                    lbl_Ref1_Em.Text = userData[10];

                    lbl_Ref2_N.Text = userData[11];
                    lbl_Ref2_PT.Text = userData[12];
                    lbl_Ref2_Org.Text = userData[13];
                    lbl_Ref2_Ph.Text = userData[14];
                    lbl_Ref2_Em.Text = userData[15];

                    lbl_Ref3_N.Text = userData[16];
                    lbl_Ref3_PT.Text = userData[17];
                    lbl_Ref3_Org.Text = userData[18];
                    lbl_Ref3_Ph.Text = userData[19];
                    lbl_Ref3_Em.Text = userData[20];

                    lbl_Ref4_N.Text = userData[21];
                    lbl_Ref4_PT.Text = userData[22];
                    lbl_Ref4_Org.Text = userData[23];
                    lbl_Ref4_Ph.Text = userData[24];
                    lbl_Ref4_Em.Text = userData[25];

                    lbl_Ref5_N.Text = userData[26];
                    lbl_Ref5_PT.Text = userData[27];
                    lbl_Ref5_Org.Text = userData[28];
                    lbl_Ref5_Ph.Text = userData[29];
                    lbl_Ref5_Em.Text = userData[30];
                }
            };

            // Print setup
            printDocument = new PrintDocument();
            printDocument.PrintPage += new PrintPageEventHandler(PrintDocument_PrintPage);
            PaperSize a4Size = new PaperSize("A4", 827, 1169);
            printDocument.DefaultPageSettings.PaperSize = a4Size;
            printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

            this.Resize += A4_References_Resize;

            // Font and color setup
            LoadFonts();
            numericUpDown_Size_Name.Value = 25;
            numericUpDown_Size_PT.Value = 12;
            numericUpDown_Size_B.Value = 10;
            numericUpDown_Size_H.Value = 12;

            button_Color_Name.BackColor = Color.Black;
            button_Color_PT.BackColor = Color.Black;
            button_Color_B.BackColor = Color.Black;
            button_Color_H.BackColor = Color.Black;
            button_IconColor.BackColor = Color.Black;

            comboBox_FontName_Name.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_Name.ValueChanged += FontSettingsChanged;
            button_Color_Name.Click += btnFontColor_Click;

            comboBox_FontName_PT.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_PT.ValueChanged += FontSettingsChanged;
            button_Color_PT.Click += btnColor_PT_Click;

            comboBox_FontName_B.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_B.ValueChanged += FontSettingsChanged;
            button_Color_B.Click += btnColor_B_Click;

            comboBox_FontName_H.SelectedIndexChanged += FontSettingsChanged;
            numericUpDown_Size_H.ValueChanged += FontSettingsChanged;
            button_Color_H.Click += btnColor_H_Click;

            button_IconColor.Click += btn_IconColor_Click;
            btn_Reset.Click += btn_Reset_Click;
            btn_RemoveIcons.Click += btn_RemoveIcons_Click;

            this.Load += (s, e) => panelA4.AutoScrollPosition = new Point(0, 0);
            this.Load += SaveOriginalControlBounds;
            this.Load += SetMaximumFormSize;

            string iconsFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Icons");
            string icon1Path = Path.Combine(iconsFolder, "icon1.png");
            string icon2Path = Path.Combine(iconsFolder, "icon2.png");
            string icon3Path = Path.Combine(iconsFolder, "icon3.png");
            string icon4Path = Path.Combine(iconsFolder, "icon4.png");

            if (File.Exists(icon1Path)) { pictureBox_Icon1.Image = Image.FromFile(icon1Path); defaultIcon1 = pictureBox_Icon1.Image; }
            if (File.Exists(icon2Path)) { pictureBox_Icon2.Image = Image.FromFile(icon2Path); defaultIcon2 = pictureBox_Icon2.Image; }
            if (File.Exists(icon3Path)) { pictureBox_Icon3.Image = Image.FromFile(icon3Path); defaultIcon3 = pictureBox_Icon3.Image; }
            if (File.Exists(icon4Path)) { pictureBox_Icon4.Image = Image.FromFile(icon4Path); defaultIcon4 = pictureBox_Icon4.Image; }

            AttachAutosaveHandlers(this);
            autosaveTimer = new Timer();
            autosaveTimer.Interval = 120000;
            autosaveTimer.Tick += AutosaveTimer_Tick;
            autosaveTimer.Start();
            this.FormClosing += A4_References_FormClosing;
        }

        // BUTTON EVENT HANDLERS

        private void pictureBox_Print_Click_1(object sender, EventArgs e)
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
            var a4_Reference_Info = new A4_Reference_Info();
            a4_Reference_Info.Size = this.Size;
            a4_Reference_Info.Show();
            this.Close();
        }

        private void A4_References_Resize(object sender, EventArgs e)
        {
            panelA4.Size = new Size(fixedWidth, fixedHeight);
            if (panelControls != null)
            {
                panelA4.Location = new Point((this.ClientSize.Width - panelA4.Width) / 2, panelControls.Bottom + 10);
            }
        }

        private void SetMaximumFormSize(object sender, EventArgs e)
        {
            if (panelControls != null && panelA4 != null)
            {
                int maxHeight = panelControls.Height + fixedHeight + SystemInformation.CaptionHeight + (SystemInformation.BorderSize.Height * 2);
                this.MaximumSize = new Size(820, maxHeight);
            }
        }

        // PRINTING METHOD (using panelA4 only)
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

        // AUTOSAVE METHODS

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
            Properties.Settings.Default.A4_References_Name = label_Name.Text;
            Properties.Settings.Default.A4_References_PositionTitle = label_PositionTitle.Text;
            Properties.Settings.Default.A4_References_Phone = label_Phone.Text;
            Properties.Settings.Default.A4_References_Email = label_Email.Text;
            Properties.Settings.Default.A4_References_Location = label_Location.Text;
            Properties.Settings.Default.A4_References_LinkedIn = label_LinkedIn.Text;

            // Save cover letter–specific settings.
            Properties.Settings.Default.A4_Ref_RefName1 = lbl_Ref1_N.Text;
            Properties.Settings.Default.A4_Ref_PositionT1 = lbl_Ref1_PT.Text;
            Properties.Settings.Default.A4_Ref_Org1 = lbl_Ref1_Org.Text;
            Properties.Settings.Default.A4_Ref_Phone1 = lbl_Ref1_Ph.Text;
            Properties.Settings.Default.A4_Ref_Email1 = lbl_Ref1_Em.Text;
            Properties.Settings.Default.A4_Ref_RefName2 = lbl_Ref2_N.Text;
            Properties.Settings.Default.A4_Ref_PositionT2 = lbl_Ref2_PT.Text;
            Properties.Settings.Default.A4_Ref_Org2 = lbl_Ref2_Org.Text;
            Properties.Settings.Default.A4_Ref_Phone2 = lbl_Ref2_Ph.Text;
            Properties.Settings.Default.A4_Ref_Email2 = lbl_Ref2_Em.Text;
            Properties.Settings.Default.A4_Ref_RefName3 = lbl_Ref3_N.Text;
            Properties.Settings.Default.A4_Ref_PositionT3 = lbl_Ref3_PT.Text;
            Properties.Settings.Default.A4_Ref_Org3 = lbl_Ref3_Org.Text;
            Properties.Settings.Default.A4_Ref_Phone3 = lbl_Ref3_Ph.Text;
            Properties.Settings.Default.A4_Ref_Email3 = lbl_Ref3_Em.Text;
            Properties.Settings.Default.A4_Ref_RefName4 = lbl_Ref4_N.Text;
            Properties.Settings.Default.A4_Ref_PositionT4 = lbl_Ref4_PT.Text;
            Properties.Settings.Default.A4_Ref_Org4 = lbl_Ref4_Org.Text;
            Properties.Settings.Default.A4_Ref_Phone4 = lbl_Ref4_Ph.Text;
            Properties.Settings.Default.A4_Ref_Email4 = lbl_Ref4_Em.Text;
            Properties.Settings.Default.A4_Ref_RefName5 = lbl_Ref5_N.Text;
            Properties.Settings.Default.A4_Ref_PositionT5 = lbl_Ref5_PT.Text;
            Properties.Settings.Default.A4_Ref_Org5 = lbl_Ref5_Org.Text;
            Properties.Settings.Default.A4_Ref_Phone5 = lbl_Ref5_Ph.Text;
            Properties.Settings.Default.A4_Ref_Email5 = lbl_Ref5_Em.Text;

            // Save font settings.
            Properties.Settings.Default.A4_References_FontName_Name = comboBox_FontName_Name.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_References_FontSize_Name = (decimal)numericUpDown_Size_Name.Value;
            Properties.Settings.Default.A4_References_Color_Name = button_Color_Name.BackColor.ToArgb();

            Properties.Settings.Default.A4_References_FontName_PT = comboBox_FontName_PT.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_References_FontSize_PT = (decimal)numericUpDown_Size_PT.Value;
            Properties.Settings.Default.A4_References_Color_PT = button_Color_PT.BackColor.ToArgb();

            Properties.Settings.Default.A4_References_FontName_Body = comboBox_FontName_B.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_References_FontSize_Body = (decimal)numericUpDown_Size_B.Value;
            Properties.Settings.Default.A4_References_Color_Body = button_Color_B.BackColor.ToArgb();

            Properties.Settings.Default.A4_References_FontName_H = comboBox_FontName_H.SelectedItem?.ToString() ?? "Microsoft Sans Serif";
            Properties.Settings.Default.A4_References_FontSize_H = (decimal)numericUpDown_Size_H.Value;
            Properties.Settings.Default.A4_References_Color_H = button_Color_H.BackColor.ToArgb();

            // Save icon file paths and icon color.
            Properties.Settings.Default.A4_References_Icon1Path = pictureBox_Icon1.Tag as string ?? "";
            Properties.Settings.Default.A4_References_Icon2Path = pictureBox_Icon2.Tag as string ?? "";
            Properties.Settings.Default.A4_References_Icon3Path = pictureBox_Icon3.Tag as string ?? "";
            Properties.Settings.Default.A4_References_Icon4Path = pictureBox_Icon4.Tag as string ?? "";
            Properties.Settings.Default.A4_References_IconColor = button_IconColor.BackColor.ToArgb();

            // Persist each icon’s Visible state:
            Properties.Settings.Default.A4_References_Icon1Visible = pictureBox_Icon1.Visible;
            Properties.Settings.Default.A4_References_Icon2Visible = pictureBox_Icon2.Visible;
            Properties.Settings.Default.A4_References_Icon3Visible = pictureBox_Icon3.Visible;
            Properties.Settings.Default.A4_References_Icon4Visible = pictureBox_Icon4.Visible;

            Properties.Settings.Default.Save();
            unsavedChanges = false;
        }

        private void LoadUserSettings()
        {
            label_Name.Text = Properties.Settings.Default.A4_References_Name;
            label_PositionTitle.Text = Properties.Settings.Default.A4_References_PositionTitle;
            label_Phone.Text = Properties.Settings.Default.A4_References_Phone;
            label_Email.Text = Properties.Settings.Default.A4_References_Email;
            label_Location.Text = Properties.Settings.Default.A4_References_Location;
            label_LinkedIn.Text = Properties.Settings.Default.A4_References_LinkedIn;
            lbl_Ref1_N.Text = Properties.Settings.Default.A4_Ref_RefName1;
            lbl_Ref1_PT.Text = Properties.Settings.Default.A4_Ref_PositionT1;
            lbl_Ref1_Org.Text = Properties.Settings.Default.A4_Ref_Org1;
            lbl_Ref1_Ph.Text = Properties.Settings.Default.A4_Ref_Phone1;
            lbl_Ref1_Em.Text = Properties.Settings.Default.A4_Ref_Email1;
            lbl_Ref2_N.Text = Properties.Settings.Default.A4_Ref_RefName2;
            lbl_Ref2_PT.Text = Properties.Settings.Default.A4_Ref_PositionT2;
            lbl_Ref2_Org.Text = Properties.Settings.Default.A4_Ref_Org2;
            lbl_Ref2_Ph.Text = Properties.Settings.Default.A4_Ref_Phone2;
            lbl_Ref2_Em.Text = Properties.Settings.Default.A4_Ref_Email2;
            lbl_Ref3_N.Text = Properties.Settings.Default.A4_Ref_RefName3;
            lbl_Ref3_PT.Text = Properties.Settings.Default.A4_Ref_PositionT3;
            lbl_Ref3_Org.Text = Properties.Settings.Default.A4_Ref_Org3;
            lbl_Ref3_Ph.Text = Properties.Settings.Default.A4_Ref_Phone3;
            lbl_Ref3_Em.Text = Properties.Settings.Default.A4_Ref_Email3;
            lbl_Ref4_N.Text = Properties.Settings.Default.A4_Ref_RefName4;
            lbl_Ref4_PT.Text = Properties.Settings.Default.A4_Ref_PositionT4;
            lbl_Ref4_Org.Text = Properties.Settings.Default.A4_Ref_Org4;
            lbl_Ref4_Ph.Text = Properties.Settings.Default.A4_Ref_Phone4;
            lbl_Ref4_Em.Text = Properties.Settings.Default.A4_Ref_Email4;
            lbl_Ref5_N.Text = Properties.Settings.Default.A4_Ref_RefName5;
            lbl_Ref5_PT.Text = Properties.Settings.Default.A4_Ref_PositionT5;
            lbl_Ref5_Org.Text = Properties.Settings.Default.A4_Ref_Org5;
            lbl_Ref5_Ph.Text = Properties.Settings.Default.A4_Ref_Phone5;
            lbl_Ref5_Em.Text = Properties.Settings.Default.A4_Ref_Email5;

            // Name
            string savedFontName = Properties.Settings.Default.A4_References_FontName_Name;
            comboBox_FontName_Name.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_Name.Value = Properties.Settings.Default.A4_References_FontSize_Name > 0
                ? Properties.Settings.Default.A4_References_FontSize_Name : 25;
            int colorVal = Properties.Settings.Default.A4_References_Color_Name;
            button_Color_Name.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Position Title
            savedFontName = Properties.Settings.Default.A4_References_FontName_PT;
            comboBox_FontName_PT.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_PT.Value = Properties.Settings.Default.A4_References_FontSize_PT > 0
                ? Properties.Settings.Default.A4_References_FontSize_PT : 12;
            colorVal = Properties.Settings.Default.A4_References_Color_PT;
            button_Color_PT.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Body
            savedFontName = Properties.Settings.Default.A4_References_FontName_Body;
            comboBox_FontName_B.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_B.Value = Properties.Settings.Default.A4_References_FontSize_Body > 0
                ? Properties.Settings.Default.A4_References_FontSize_Body : 10;
            colorVal = Properties.Settings.Default.A4_References_Color_Body;
            button_Color_B.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Headings
            savedFontName = Properties.Settings.Default.A4_References_FontName_H;
            comboBox_FontName_H.SelectedItem = !string.IsNullOrEmpty(savedFontName)
                ? savedFontName : "Microsoft Sans Serif";
            numericUpDown_Size_H.Value = Properties.Settings.Default.A4_References_FontSize_H > 0
                ? Properties.Settings.Default.A4_References_FontSize_H : 12;
            colorVal = Properties.Settings.Default.A4_References_Color_H;
            button_Color_H.BackColor = (colorVal == 0) ? Color.Black : Color.FromArgb(colorVal);

            // Icon Color
            int iconColorVal = Properties.Settings.Default.A4_References_IconColor;
            button_IconColor.BackColor = (iconColorVal == 0) ? Color.Black : Color.FromArgb(iconColorVal);

            // Match reset style for icon color button
            button_IconColor.FlatStyle = FlatStyle.Flat;
            button_IconColor.ForeColor = Color.White;
            button_IconColor.FlatAppearance.BorderColor = Color.White;
            button_IconColor.UseVisualStyleBackColor = false;

            // Apply icons and fonts
            ApplyIconColor(button_IconColor.BackColor);
            ApplyFontChanges();

            string icon1Path = Properties.Settings.Default.A4_References_Icon1Path;
            string icon2Path = Properties.Settings.Default.A4_References_Icon2Path;
            string icon3Path = Properties.Settings.Default.A4_References_Icon3Path;
            string icon4Path = Properties.Settings.Default.A4_References_Icon4Path;

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
            pictureBox_Icon1.Visible = Properties.Settings.Default.A4_References_Icon1Visible;
            pictureBox_Icon2.Visible = Properties.Settings.Default.A4_References_Icon2Visible;
            pictureBox_Icon3.Visible = Properties.Settings.Default.A4_References_Icon3Visible;
            pictureBox_Icon4.Visible = Properties.Settings.Default.A4_References_Icon4Visible;
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
            foreach (Control ctrl in panelA4.Controls)
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
                comboBox_FontName_H.Items.Add(font.Name);
            }
            comboBox_FontName_Name.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_PT.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_B.SelectedItem = "Microsoft Sans Serif";
            comboBox_FontName_H.SelectedItem = "Microsoft Sans Serif";
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
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Name"))
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
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "PositionTitle"))
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
                foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
                {
                    lbl.Font = newFont;
                    lbl.ForeColor = color;
                }
            }

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
                    AutosaveUserSettings(); // Save immediately.
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

            comboBox_FontName_H.SelectedItem = "Microsoft Sans Serif";
            numericUpDown_Size_H.Value = 12;
            button_Color_H.BackColor = Color.Black;

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
            foreach (Label lbl in GetAllControls(panelA4).OfType<Label>().Where(x => x.Tag?.ToString() == "Body"))
            {
                lbl.Font = new Font("Microsoft Sans Serif", 10, FontStyle.Regular);
                lbl.ForeColor = Color.Black;
            }
        }

        private void btn_ResetIcons_Click(object sender, EventArgs e)
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

        private void A4_References_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (unsavedChanges)
                AutosaveUserSettings();
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

        private void btn_RemoveIcons_Click(object sender, EventArgs e)
        {
            var icons = GetAllControls(panelA4).OfType<PictureBox>()
                        .Where(x => x.Name.StartsWith("pictureBox_Icon"))
                        .ToList();
            MessageBox.Show("Found " + icons.Count + " icons.");
            foreach (PictureBox pic in icons)
                pic.Visible = false;
            unsavedChanges = true;
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

        private void button_Save_Click(object sender, EventArgs e)
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

        // LOAD AND RESIZE EVENT HANDLERS

        private void A4_CoverLetter_Load(object sender, EventArgs e)
        {
            if (panelControls != null)
            {
                panelA4.Location = new Point((this.ClientSize.Width - panelA4.Width) / 2, panelControls.Bottom + 10);
            }
        }       
    }
}




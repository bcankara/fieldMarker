using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.Drawing.Drawing2D;

namespace fieldMarker
{
    [ComVisible(true)]
    public class CustomRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private Word.WdColorIndex selectedColor = Word.WdColorIndex.wdNoHighlight;
        private Dictionary<string, Word.WdColorIndex> colorMap = new Dictionary<string, Word.WdColorIndex>()
        {
            { "btnNoColor", Word.WdColorIndex.wdNoHighlight },
            { "btnYellow", Word.WdColorIndex.wdYellow },
            { "btnGreen", Word.WdColorIndex.wdBrightGreen },
            { "btnTurquoise", Word.WdColorIndex.wdTurquoise },
            { "btnPink", Word.WdColorIndex.wdPink },
            { "btnRed", Word.WdColorIndex.wdRed },
            { "btnBlue", Word.WdColorIndex.wdBlue },
            { "btnDarkBlue", Word.WdColorIndex.wdDarkBlue },
            { "btnTeal", Word.WdColorIndex.wdTeal },
            { "btnGray", Word.WdColorIndex.wdGray25 }
        };

        public CustomRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string resourceName = "fieldMarker.CustomRibbon.xml";
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnColorButtonClick(Office.IRibbonControl control)
        {
            try
            {
                ApplyHighlight();
            }
            catch (Exception ex)
            {
                ShowNotification($"Error applying highlight: {ex.Message}", "Error", true);
            }
        }

        public void OnColorSelected(Office.IRibbonControl control)
        {
            try
            {
                if (control != null && control.Id != null && colorMap.ContainsKey(control.Id))
                {
                    selectedColor = colorMap[control.Id];
                    ApplyHighlight();
                }
            }
            catch (Exception ex)
            {
                ShowNotification($"Error selecting color: {ex.Message}", "Error", true);
            }
        }

        public void OnInfoButtonClick(Office.IRibbonControl control)
        {
            var result = MessageBox.Show(
                "Developer: Burak Can KARA\nEmail: burakcankara@outlook.com\nGitHub: github.com/bcankara\n\nClick OK to visit GitHub profile",
                "Field Marker Add-In",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (result == DialogResult.OK)
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "https://github.com/bcankara",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowNotification($"Error opening link: {ex.Message}", "Error", true);
                }
            }
        }

        #endregion

        #region Image Getters

        public Bitmap GetInfoImage(Office.IRibbonControl control)
        {
            return Properties.Resources.info;
        }

        public Bitmap GetColorPickerImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_picker;
        }

        public Bitmap GetNoColorImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_nocolor;
        }

        public Bitmap GetYellowImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_yellow;
        }

        public Bitmap GetGreenImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_green;
        }

        public Bitmap GetTurquoiseImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_turquoise;
        }

        public Bitmap GetPinkImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_pink;
        }

        public Bitmap GetRedImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_red;
        }

        public Bitmap GetBlueImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_blue;
        }

        public Bitmap GetDarkBlueImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_darkblue;
        }

        public Bitmap GetTealImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_teal;
        }

        public Bitmap GetGrayImage(Office.IRibbonControl control)
        {
            return Properties.Resources.color_gray;
        }

        #endregion

        #region Helper Methods

        private void ShowNotification(string message, string title, bool isError = false)
        {
            NotificationForm.Show(message, title, isError);
        }

        private class ModernProgressForm : Form
        {
            private readonly ProgressBar progressBar;
            private readonly Label statusLabel;
            private readonly Label percentLabel;
            private readonly Label completionLabel;
            private readonly Button okButton;

            public ModernProgressForm()
            {
                this.Text = "Field Marker Progress";
                this.Width = 400;
                this.Height = 200;
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = Color.White;
                this.TopMost = true;
                
                // Form gölgesi ve yuvarlak köşeler
                this.Paint += (s, e) =>
                {
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var path = new GraphicsPath())
                    {
                        path.AddRoundedRectangle(this.ClientRectangle, 10);
                        e.Graphics.FillPath(new SolidBrush(this.BackColor), path);
                    }
                };

                // Form sürükleme için mouse events
                bool isDragging = false;
                Point dragStartPoint = Point.Empty;

                this.MouseDown += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                    {
                        isDragging = true;
                        dragStartPoint = new Point(e.X, e.Y);
                    }
                };

                this.MouseMove += (s, e) =>
                {
                    if (isDragging)
                    {
                        Point p = PointToScreen(e.Location);
                        Location = new Point(p.X - dragStartPoint.X, p.Y - dragStartPoint.Y);
                    }
                };

                this.MouseUp += (s, e) =>
                {
                    isDragging = false;
                };

                // Başlık etiketi
                var titleLabel = new Label
                {
                    Text = "Processing Fields",
                    Font = new Font("Segoe UI", 12, FontStyle.Bold),
                    ForeColor = Color.FromArgb(64, 64, 64),
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Width = 360,
                    Height = 30,
                    Location = new Point(20, 15)
                };

                // Progress bar
                progressBar = new ProgressBar
                {
                    Width = 360,
                    Height = 23,
                    Location = new Point(20, 55),
                    Style = ProgressBarStyle.Continuous
                };

                // Durum etiketi
                statusLabel = new Label
                {
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Width = 280,
                    Height = 23,
                    Location = new Point(20, 90),
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.FromArgb(64, 64, 64)
                };

                // Yüzde etiketi
                percentLabel = new Label
                {
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleRight,
                    Width = 80,
                    Height = 23,
                    Location = new Point(300, 90),
                    Font = new Font("Segoe UI", 9, FontStyle.Bold),
                    ForeColor = Color.FromArgb(64, 64, 64)
                };

                // Tamamlanma mesajı etiketi
                completionLabel = new Label
                {
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Width = 360,
                    Height = 30,
                    Location = new Point(20, 120),
                    Font = new Font("Segoe UI", 9),
                    ForeColor = Color.FromArgb(64, 64, 64),
                    Visible = false
                };

                // OK Butonu
                okButton = new Button
                {
                    Text = "✓ OK",
                    Width = 120,
                    Height = 35,
                    Location = new Point((this.Width - 120) / 2, 155),
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 11, FontStyle.Bold),
                    Visible = false,
                    Name = "okButton",
                    BackColor = Color.FromArgb(0, 120, 212),
                    ForeColor = Color.White,
                    Cursor = Cursors.Hand
                };

                // Buton köşelerini yuvarla
                okButton.Paint += (s, e) =>
                {
                    var buttonPath = new GraphicsPath();
                    var rect = new Rectangle(0, 0, okButton.Width, okButton.Height);
                    buttonPath.AddRoundedRectangle(rect, 5);
                    okButton.Region = new Region(buttonPath);
                };

                // Buton efektleri
                okButton.MouseEnter += (s, e) => okButton.BackColor = Color.FromArgb(0, 102, 204);
                okButton.MouseLeave += (s, e) => okButton.BackColor = Color.FromArgb(0, 120, 212);
                okButton.Click += (s, e) => this.Close();

                this.Controls.AddRange(new Control[] { titleLabel, progressBar, statusLabel, percentLabel, completionLabel, okButton });

                // ESC tuşu ile kapatma
                this.KeyPreview = true;
                this.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Escape && okButton.Visible)
                    {
                        this.Close();
                    }
                };

                // Enter tuşu ile kapatma
                this.AcceptButton = okButton;
            }

            public void UpdateProgress(int current, int total)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => UpdateProgress(current, total)));
                    return;
                }

                int percentage = (int)((current / (double)total) * 100);
                progressBar.Value = percentage;
                statusLabel.Text = $"Processing field {current} of {total}";
                percentLabel.Text = $"{percentage}%";
                Application.DoEvents();
            }

            public void ShowCompletion(string message)
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => ShowCompletion(message)));
                    return;
                }

                completionLabel.Text = message;
                completionLabel.Visible = true;
                progressBar.Value = 100;
                
                okButton.Visible = true;
                okButton.Focus();

                statusLabel.Text = "Operation completed successfully";
                percentLabel.Text = "100%";
                
                // Progress bar'ı yeşil yap
                progressBar.ForeColor = Color.FromArgb(40, 167, 69);
                
                this.TopMost = true;
                Application.DoEvents();

                // Başlık metnini güncelle
                var titleLabel = this.Controls.OfType<Label>().FirstOrDefault(l => l.Font.Size == 12);
                if (titleLabel != null)
                {
                    titleLabel.Text = "Operation Complete";
                }
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    const int CS_DROPSHADOW = 0x20000;
                    CreateParams cp = base.CreateParams;
                    cp.ClassStyle |= CS_DROPSHADOW;
                    return cp;
                }
            }
        }

        private void ApplyHighlight()
        {
            ModernProgressForm progressForm = null;
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app == null || app.ActiveDocument == null)
                {
                    MessageBox.Show("Word document is not available.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var document = app.ActiveDocument;
                var fields = document.Fields.Cast<Word.Field>().ToList();
                int totalFields = fields.Count;

                if (totalFields == 0)
                {
                    MessageBox.Show("No fields found in the document.", "No Fields Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                progressForm = new ModernProgressForm();
                progressForm.Show();
                int fieldCount = 0;

                // Screen updating'i kapat
                app.ScreenUpdating = false;

                foreach (Word.Field field in fields)
                {
                    try
                    {
                        if (field?.Result != null)
                        {
                            // Alan kodunun tüm aralığını seç
                            var range = field.Result;
                            range.HighlightColorIndex = selectedColor;
                            fieldCount++;

                            // İlerleme çubuğunu güncelle
                            progressForm.UpdateProgress(fieldCount, totalFields);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }

                // Screen updating'i geri aç
                app.ScreenUpdating = true;

                if (fieldCount > 0)
                {
                    string colorName = colorMap.FirstOrDefault(x => x.Value == selectedColor).Key?.Replace("btn", "");
                    string message = $"{fieldCount} fields {(selectedColor == Word.WdColorIndex.wdNoHighlight ? "cleared" : $"highlighted with {colorName} color")}.";
                    progressForm.ShowCompletion(message);
                }
                else
                {
                    progressForm.Close();
                }
            }
            catch (Exception ex)
            {
                if (progressForm != null && !progressForm.IsDisposed)
                {
                    progressForm.Close();
                }
                MessageBox.Show($"Error processing fields: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Screen updating'in açık olduğundan emin ol
                if (Globals.ThisAddIn.Application != null)
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
            }
        }

        #endregion
    }

    public static class GraphicsExtensions
    {
        public static void AddRoundedRectangle(this GraphicsPath path, Rectangle bounds, int radius)
        {
            path.AddArc(bounds.X, bounds.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(bounds.Right - radius * 2, bounds.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(bounds.Right - radius * 2, bounds.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
        }
    }
} 
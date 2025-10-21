using System;
using System.Drawing;
using System.Windows.Forms;
using Grasshopper;

namespace GhcETABSAPI
{
    internal static class UiHelpers
    {
        private static readonly object _syncRoot = new object();
        private static ProgressWindow _progressWindow;

        internal static void ShowProgressBar(string title, string initialStatus, int maximum)
        {
            ShowDualProgressBar(title, string.Empty, 0, initialStatus, maximum);
        }

        internal static void ShowDualProgressBar(
            string title,
            string excelStatus,
            int excelMaximum,
            string assignmentStatus,
            int assignmentMaximum)
        {
            if (excelMaximum < 0)
            {
                excelMaximum = 0;
            }

            if (assignmentMaximum < 0)
            {
                assignmentMaximum = 0;
            }

            RunOnUiThread(() =>
            {
                lock (_syncRoot)
                {
                    EnsureWindow();

                    _progressWindow.Text = string.IsNullOrWhiteSpace(title) ? "Progress" : title;
                    _progressWindow.InitializeDual(excelStatus, excelMaximum, assignmentStatus, assignmentMaximum);

                    if (!_progressWindow.Visible)
                    {
                        Form owner = Grasshopper.Instances.DocumentEditor;
                        if (owner != null && !owner.IsDisposed)
                        {
                            PositionRelativeToOwner(owner, _progressWindow);
                            _progressWindow.Show(owner);
                        }
                        else
                        {
                            _progressWindow.StartPosition = FormStartPosition.CenterScreen;
                            _progressWindow.Show();
                        }
                    }
                    else
                    {
                        _progressWindow.BringToFront();
                    }
                }
            });
        }

        internal static void UpdateProgressBar(int value, int maximum, string status)
        {
            UpdateAssignmentProgressBar(value, maximum, status);
        }

        internal static void UpdateExcelProgressBar(int value, int maximum, string status)
        {
            if (maximum < 0)
            {
                maximum = 0;
            }

            RunOnUiThread(() =>
            {
                lock (_syncRoot)
                {
                    if (_progressWindow == null || _progressWindow.IsDisposed)
                    {
                        return;
                    }

                    _progressWindow.UpdateExcelProgress(value, maximum, status);
                }
            });
        }

        internal static void UpdateAssignmentProgressBar(int value, int maximum, string status)
        {
            if (maximum < 0)
            {
                maximum = 0;
            }

            RunOnUiThread(() =>
            {
                lock (_syncRoot)
                {
                    if (_progressWindow == null || _progressWindow.IsDisposed)
                    {
                        return;
                    }

                    _progressWindow.UpdateAssignmentProgress(value, maximum, status);
                }
            });
        }

        internal static void CloseProgressBar()
        {
            RunOnUiThread(() =>
            {
                lock (_syncRoot)
                {
                    if (_progressWindow != null)
                    {
                        try
                        {
                            if (_progressWindow.Visible)
                            {
                                _progressWindow.Hide();
                            }
                            _progressWindow.Close();
                        }
                        catch
                        {
                            // ignored
                        }
                        finally
                        {
                            _progressWindow.Dispose();
                            _progressWindow = null;
                        }
                    }
                }
            });
        }

        private static void EnsureWindow()
        {
            if (_progressWindow == null || _progressWindow.IsDisposed)
            {
                _progressWindow = new ProgressWindow();
            }
        }

        private static void PositionRelativeToOwner(Form owner, Form window)
        {
            try
            {
                Rectangle ownerBounds = owner.Bounds;
                int x = ownerBounds.Left + Math.Max(0, (ownerBounds.Width - window.Width) / 2);
                int y = ownerBounds.Top + Math.Max(0, (ownerBounds.Height - window.Height) / 2);
                window.StartPosition = FormStartPosition.Manual;
                window.Location = new Point(x, y);
            }
            catch
            {
                window.StartPosition = FormStartPosition.CenterScreen;
            }
        }

        private static void RunOnUiThread(Action action)
        {
            if (action == null)
            {
                return;
            }

            Form editor = Grasshopper.Instances.DocumentEditor;
            if (editor != null && !editor.IsDisposed)
            {
                if (editor.InvokeRequired)
                {
                    editor.BeginInvoke(action);
                }
                else
                {
                    action();
                }
                return;
            }

            Control canvas = Grasshopper.Instances.ActiveCanvas;
            if (canvas != null && !canvas.IsDisposed)
            {
                if (canvas.InvokeRequired)
                {
                    canvas.BeginInvoke(action);
                }
                else
                {
                    action();
                }
                return;
            }

            action();
        }

        private sealed class ProgressWindow : Form
        {
            private readonly Label _excelStatusLabel;
            private readonly ProgressBar _excelProgressBar;
            private readonly Label _assignmentStatusLabel;
            private readonly ProgressBar _assignmentProgressBar;

            internal ProgressWindow()
            {
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                ControlBox = false;
                TopMost = false;
                Size = new Size(380, 170);
                Padding = new Padding(12);

                TableLayoutPanel layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 4,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink
                };

                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                _excelStatusLabel = new Label
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(0, 0, 0, 4)
                };

                _excelProgressBar = new ProgressBar
                {
                    Dock = DockStyle.Fill,
                    Height = 20,
                    Style = ProgressBarStyle.Continuous,
                    Minimum = 0,
                    Maximum = 1,
                    Margin = new Padding(0, 0, 0, 12)
                };

                _assignmentStatusLabel = new Label
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(0, 0, 0, 4)
                };

                _assignmentProgressBar = new ProgressBar
                {
                    Dock = DockStyle.Fill,
                    Height = 20,
                    Style = ProgressBarStyle.Continuous,
                    Minimum = 0,
                    Maximum = 1,
                    Margin = new Padding(0)
                };

                layout.Controls.Add(_excelStatusLabel, 0, 0);
                layout.Controls.Add(_excelProgressBar, 0, 1);
                layout.Controls.Add(_assignmentStatusLabel, 0, 2);
                layout.Controls.Add(_assignmentProgressBar, 0, 3);

                Controls.Add(layout);

                HideExcelSection();
                HideAssignmentSection();
            }

            internal void InitializeDual(string excelStatus, int excelMaximum, string assignmentStatus, int assignmentMaximum)
            {
                UpdateExcelProgress(0, excelMaximum, excelStatus);
                UpdateAssignmentProgress(0, assignmentMaximum, assignmentStatus);
            }

            internal void UpdateExcelProgress(int value, int maximum, string status)
            {
                UpdateSection(_excelStatusLabel, _excelProgressBar, value, maximum, status, HideExcelSection);
            }

            internal void UpdateAssignmentProgress(int value, int maximum, string status)
            {
                UpdateSection(_assignmentStatusLabel, _assignmentProgressBar, value, maximum, status, HideAssignmentSection);
            }

            private void UpdateSection(
                Label statusLabel,
                ProgressBar progressBar,
                int value,
                int maximum,
                string status,
                Action hideAction)
            {
                bool shouldShow = !string.IsNullOrWhiteSpace(status) || maximum > 0;
                if (!shouldShow)
                {
                    hideAction();
                    return;
                }

                statusLabel.Visible = true;
                progressBar.Visible = true;

                int safeMaximum = Math.Max(1, maximum);
                if (progressBar.Maximum != safeMaximum)
                {
                    progressBar.Maximum = safeMaximum;
                }

                int clamped = Math.Max(progressBar.Minimum, Math.Min(value, safeMaximum));
                try
                {
                    progressBar.Value = clamped;
                }
                catch
                {
                    progressBar.Value = progressBar.Minimum;
                }

                statusLabel.Text = string.IsNullOrWhiteSpace(status) ? string.Empty : status;
            }

            private void HideExcelSection()
            {
                _excelStatusLabel.Visible = false;
                _excelStatusLabel.Text = string.Empty;
                _excelProgressBar.Visible = false;
                _excelProgressBar.Value = 0;
            }

            private void HideAssignmentSection()
            {
                _assignmentStatusLabel.Visible = false;
                _assignmentStatusLabel.Text = string.Empty;
                _assignmentProgressBar.Visible = false;
                _assignmentProgressBar.Value = 0;
            }
        }
    }
}

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
            if (maximum < 0)
            {
                maximum = 0;
            }

            RunOnUiThread(() =>
            {
                lock (_syncRoot)
                {
                    EnsureWindow();

                    _progressWindow.Text = string.IsNullOrWhiteSpace(title) ? "Progress" : title;
                    _progressWindow.UpdateProgress(0, maximum, initialStatus);

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

                    _progressWindow.UpdateProgress(value, maximum, status);
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
            private readonly Label _statusLabel;
            private readonly ProgressBar _progressBar;

            internal ProgressWindow()
            {
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowInTaskbar = false;
                ControlBox = false;
                TopMost = false;
                Size = new Size(360, 120);
                Padding = new Padding(12);

                _statusLabel = new Label
                {
                    Dock = DockStyle.Top,
                    Height = 40,
                    AutoSize = false,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Text = "Working..."
                };

                _progressBar = new ProgressBar
                {
                    Dock = DockStyle.Top,
                    Height = 24,
                    Style = ProgressBarStyle.Continuous,
                    Minimum = 0,
                    Maximum = 1
                };

                Controls.Add(_progressBar);
                Controls.Add(_statusLabel);
            }

            internal void UpdateProgress(int currentValue, int maximum, string status)
            {
                int displayMax = Math.Max(1, maximum);
                if (_progressBar.Maximum != displayMax)
                {
                    _progressBar.Maximum = displayMax;
                }

                int clamped = Math.Max(_progressBar.Minimum, Math.Min(currentValue, displayMax));
                try
                {
                    _progressBar.Value = clamped;
                }
                catch
                {
                    _progressBar.Value = _progressBar.Minimum;
                }

                if (!string.IsNullOrWhiteSpace(status))
                {
                    _statusLabel.Text = status;
                }
                else if (string.IsNullOrWhiteSpace(_statusLabel.Text))
                {
                    _statusLabel.Text = "Working...";
                }
            }
        }
    }
}

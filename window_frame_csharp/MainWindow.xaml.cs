using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;

namespace WindowFramer
{
    public partial class MainWindow : Window
    {
        private const int FRAME_THICKNESS = 4;
        private IntPtr _targetHwnd = IntPtr.Zero;
        private CancellationTokenSource? _cts;
        private List<Window> _frameWindows = new();
        private SolidColorBrush _selectedBrush = new(Colors.Red);

        public MainWindow()
        {
            InitializeComponent();
            ColorComboBox.ItemsSource = new Dictionary<string, Color>
            {
                {"Red", Colors.Red}, {"Blue", Colors.Blue}, {"Green", Colors.Green},
                {"Yellow", Colors.Yellow}, {"Cyan", Colors.Cyan}, {"Magenta", Colors.Magenta},
                {"Black", Colors.Black}, {"White", Colors.White}
            };
            ColorComboBox.DisplayMemberPath = "Key";
            ColorComboBox.SelectedValuePath = "Value";
            ColorComboBox.SelectedIndex = 0;
        }

        private async void SelectWindow_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = Visibility.Collapsed;
            await Task.Delay(500); // Give time for window to hide

            // Wait for user to click on the target window
            while (true)
            {
                if ((NativeMethods.GetAsyncKeyState(0x01) & 0x8000) != 0) // VK_LBUTTON
                {
                    NativeMethods.GetCursorPos(out var point);
                    var hwnd = NativeMethods.WindowFromPoint(point);
                    _targetHwnd = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
                    break;
                }
                await Task.Delay(10);
            }
            
            this.Visibility = Visibility.Visible;

            var windowTitle = NativeMethods.GetWindowText(_targetHwnd);
            Debug.WriteLine($"Selected window: '{windowTitle}' (HWND: {_targetHwnd})");

            StartTracking();
        }

        private void StopFraming_Click(object sender, RoutedEventArgs e)
        {
            StopTracking();
        }

        private void ColorComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ColorComboBox.SelectedValue is Color color)
            {
                _selectedBrush = new SolidColorBrush(color);
                if (_cts != null && !_cts.IsCancellationRequested)
                {
                    // Re-create frames with the new color
                    Dispatcher.Invoke(() =>
                    {
                        CloseFrameWindows();
                        CreateFrameWindows();
                    });
                }
            }
        }

        private void StartTracking()
        {
            if (_cts != null && !_cts.IsCancellationRequested)
            {
                _cts.Cancel();
            }
            _cts = new CancellationTokenSource();
            var token = _cts.Token;

            Task.Run(() => TrackWindowLoop(token), token);
        }

        private void StopTracking()
        {
            _cts?.Cancel();
            Dispatcher.Invoke(CloseFrameWindows);
        }

        private void CreateFrameWindows()
        {
            for (int i = 0; i < 4; i++)
            {
                var frame = new Window
                {
                    WindowStyle = WindowStyle.None,
                    ResizeMode = ResizeMode.NoResize,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent,
                    ShowInTaskbar = false,
                    Content = new Rectangle { Fill = _selectedBrush }
                };
                var helper = new WindowInteropHelper(frame);
                helper.EnsureHandle();
                // Make the frame window always on top
                NativeMethods.SetWindowPos(helper.Handle, -1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040); // HWND_TOPMOST, SWP_NOMOVE, SWP_NOSIZE, SWP_NOACTIVATE
                frame.Show();
                _frameWindows.Add(frame);
            }
        }

        private void CloseFrameWindows()
        {
            foreach (var frame in _frameWindows)
            {
                frame.Close();
            }
            _frameWindows.Clear();
        }

        private void TrackWindowLoop(CancellationToken token)
        {
            Dispatcher.Invoke(CreateFrameWindows);

            while (!token.IsCancellationRequested)
            {
                if (!NativeMethods.IsWindow(_targetHwnd))
                {
                    Dispatcher.Invoke(StopTracking);
                    break;
                }

                // Use DwmGetWindowAttribute to get the actual window bounds, excluding shadow.
                // Fall back to GetWindowRect on failure.
                var result = NativeMethods.DwmGetWindowAttribute(_targetHwnd, NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS, out var rect, Marshal.SizeOf(typeof(NativeMethods.RECT)));
                if (result != 0) // 0 is S_OK. On failure, use the old method.
                {
                    NativeMethods.GetWindowRect(_targetHwnd, out rect);
                }

                Dispatcher.Invoke(() =>
                {
                    // Top
                    _frameWindows[0].Left = rect.Left;
                    _frameWindows[0].Top = rect.Top;
                    _frameWindows[0].Width = rect.Right - rect.Left;
                    _frameWindows[0].Height = FRAME_THICKNESS;

                    // Bottom
                    _frameWindows[1].Left = rect.Left;
                    _frameWindows[1].Top = rect.Bottom - FRAME_THICKNESS;
                    _frameWindows[1].Width = rect.Right - rect.Left;
                    _frameWindows[1].Height = FRAME_THICKNESS;

                    // Left
                    _frameWindows[2].Left = rect.Left;
                    _frameWindows[2].Top = rect.Top;
                    _frameWindows[2].Width = FRAME_THICKNESS;
                    _frameWindows[2].Height = rect.Bottom - rect.Top;

                    // Right
                    _frameWindows[3].Left = rect.Right - FRAME_THICKNESS;
                    _frameWindows[3].Top = rect.Top;
                    _frameWindows[3].Width = FRAME_THICKNESS;
                    _frameWindows[3].Height = rect.Bottom - rect.Top;
                });

                Thread.Sleep(16); // ~60 FPS
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            StopTracking();
            base.OnClosed(e);
        }
    }
}
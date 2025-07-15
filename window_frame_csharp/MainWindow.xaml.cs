using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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

            var sb = new StringBuilder(256);
            NativeMethods.GetWindowText(_targetHwnd, sb, sb.Capacity);
            var windowTitle = sb.ToString();
            
            // Get initial window rect for debugging
            NativeMethods.RECT initialRect;
            var result = NativeMethods.DwmGetWindowAttribute(_targetHwnd, NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS, out initialRect, Marshal.SizeOf(typeof(NativeMethods.RECT)));
            if (result != 0)
            {
                NativeMethods.GetWindowRect(_targetHwnd, out initialRect);
            }
            
            Debug.WriteLine($"Selected window: '{windowTitle}' (HWND: {_targetHwnd})");
            Debug.WriteLine($"Initial window rect: Left={initialRect.Left}, Top={initialRect.Top}, Right={initialRect.Right}, Bottom={initialRect.Bottom}");
            Debug.WriteLine($"Initial window size: Width={initialRect.Right - initialRect.Left}, Height={initialRect.Bottom - initialRect.Top}");

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
                
                // Set extended window style to be a tool window (doesn't steal focus)
                var exStyle = NativeMethods.GetWindowLong(helper.Handle, NativeMethods.GWL_EXSTYLE);
                NativeMethods.SetWindowLong(helper.Handle, NativeMethods.GWL_EXSTYLE, 
                    exStyle | NativeMethods.WS_EX_TOOLWINDOW | NativeMethods.WS_EX_NOACTIVATE);
                
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

                // Get window placement to check if maximized
                var placement = new NativeMethods.WINDOWPLACEMENT();
                placement.length = Marshal.SizeOf(placement);
                NativeMethods.GetWindowPlacement(_targetHwnd, ref placement);
                
                // Get the actual client area and window rectangle to calculate exact borders
                NativeMethods.RECT windowRect;
                NativeMethods.RECT clientRect;
                
                if (!NativeMethods.GetWindowRect(_targetHwnd, out windowRect))
                {
                    Thread.Sleep(16);
                    continue;
                }
                
                if (!NativeMethods.GetClientRect(_targetHwnd, out clientRect))
                {
                    Thread.Sleep(16);
                    continue;
                }
                
                // Get monitor information for screen bounds
                IntPtr monitor = NativeMethods.MonitorFromWindow(_targetHwnd, NativeMethods.MONITOR_DEFAULTTONEAREST);
                var monitorInfo = new NativeMethods.MONITORINFO();
                monitorInfo.cbSize = Marshal.SizeOf(monitorInfo);
                NativeMethods.GetMonitorInfo(monitor, ref monitorInfo);
                
                // Convert client rect to screen coordinates
                var clientTopLeft = new NativeMethods.POINT { X = clientRect.Left, Y = clientRect.Top };
                var clientBottomRight = new NativeMethods.POINT { X = clientRect.Right, Y = clientRect.Bottom };
                
                NativeMethods.ClientToScreen(_targetHwnd, ref clientTopLeft);
                NativeMethods.ClientToScreen(_targetHwnd, ref clientBottomRight);
                
                // Calculate the exact window border sizes
                int leftBorder = clientTopLeft.X - windowRect.Left;
                int topBorder = clientTopLeft.Y - windowRect.Top;
                int rightBorder = windowRect.Right - clientBottomRight.X;
                int bottomBorder = windowRect.Bottom - clientBottomRight.Y;
                
                // The actual visible window contour (excluding shadow)
                NativeMethods.RECT visibleRect = new NativeMethods.RECT
                {
                    Left = windowRect.Left + Math.Max(0, leftBorder - 1),
                    Top = windowRect.Top + Math.Max(0, topBorder - 1),
                    Right = windowRect.Right - Math.Max(0, rightBorder - 1),
                    Bottom = windowRect.Bottom - Math.Max(0, bottomBorder - 1)
                };

                Dispatcher.Invoke(() =>
                {
                    // Debug output for exact border calculations
                    Debug.WriteLine($"Window rect: L={windowRect.Left}, T={windowRect.Top}, R={windowRect.Right}, B={windowRect.Bottom}");
                    Debug.WriteLine($"Client rect (screen): L={clientTopLeft.X}, T={clientTopLeft.Y}, R={clientBottomRight.X}, B={clientBottomRight.Y}");
                    Debug.WriteLine($"Borders: L={leftBorder}, T={topBorder}, R={rightBorder}, B={bottomBorder}");
                    Debug.WriteLine($"Visible rect: L={visibleRect.Left}, T={visibleRect.Top}, R={visibleRect.Right}, B={visibleRect.Bottom}");
                    
                    // Basic validation
                    if (visibleRect.Right <= visibleRect.Left || visibleRect.Bottom <= visibleRect.Top)
                    {
                        Debug.WriteLine("Invalid visible rectangle, skipping");
                        return;
                    }

                    // Convert Win32 coordinates to WPF coordinates
                    var source = PresentationSource.FromVisual(this);
                    double dpiScaleX = 1.0;
                    double dpiScaleY = 1.0;
                    
                    if (source?.CompositionTarget != null)
                    {
                        dpiScaleX = source.CompositionTarget.TransformToDevice.M11;
                        dpiScaleY = source.CompositionTarget.TransformToDevice.M22;
                    }

                    // Convert to WPF coordinates
                    double left = visibleRect.Left / dpiScaleX;
                    double top = visibleRect.Top / dpiScaleY;
                    double width = (visibleRect.Right - visibleRect.Left) / dpiScaleX;
                    double height = (visibleRect.Bottom - visibleRect.Top) / dpiScaleY;
                    
                    Debug.WriteLine($"WPF coords: L={left:F1}, T={top:F1}, W={width:F1}, H={height:F1}");
                    Debug.WriteLine($"Monitor work area: L={monitorInfo.rcWork.Left}, T={monitorInfo.rcWork.Top}, R={monitorInfo.rcWork.Right}, B={monitorInfo.rcWork.Bottom}");
                    Debug.WriteLine($"Window maximized: {placement.showCmd == NativeMethods.SW_SHOWMAXIMIZED}");

                    // Check if window is at screen edges (maximized or snapped)
                    bool atLeftEdge = Math.Abs(visibleRect.Left - monitorInfo.rcWork.Left) < 5;
                    bool atTopEdge = Math.Abs(visibleRect.Top - monitorInfo.rcWork.Top) < 5;
                    bool atRightEdge = Math.Abs(visibleRect.Right - monitorInfo.rcWork.Right) < 5;
                    bool atBottomEdge = Math.Abs(visibleRect.Bottom - monitorInfo.rcWork.Bottom) < 5;
                    
                    Debug.WriteLine($"At edges: L={atLeftEdge}, T={atTopEdge}, R={atRightEdge}, B={atBottomEdge}");

                    // Position frame windows with adjustments for screen edges
                    // Top frame - if at top edge, position inside the window
                    if (atTopEdge)
                    {
                        _frameWindows[0].Left = left;
                        _frameWindows[0].Top = top;
                        _frameWindows[0].Width = width;
                        _frameWindows[0].Height = FRAME_THICKNESS;
                    }
                    else
                    {
                        _frameWindows[0].Left = left - FRAME_THICKNESS;
                        _frameWindows[0].Top = top - FRAME_THICKNESS;
                        _frameWindows[0].Width = width + (2 * FRAME_THICKNESS);
                        _frameWindows[0].Height = FRAME_THICKNESS;
                    }
                    
                    // Bottom frame - if at bottom edge, position inside the window
                    if (atBottomEdge)
                    {
                        _frameWindows[1].Left = left;
                        _frameWindows[1].Top = top + height - FRAME_THICKNESS;
                        _frameWindows[1].Width = width;
                        _frameWindows[1].Height = FRAME_THICKNESS;
                    }
                    else
                    {
                        _frameWindows[1].Left = left - FRAME_THICKNESS;
                        _frameWindows[1].Top = top + height;
                        _frameWindows[1].Width = width + (2 * FRAME_THICKNESS);
                        _frameWindows[1].Height = FRAME_THICKNESS;
                    }
                    
                    // Left frame - if at left edge, position inside the window
                    if (atLeftEdge)
                    {
                        _frameWindows[2].Left = left;
                        _frameWindows[2].Top = atTopEdge ? top + FRAME_THICKNESS : top;
                        _frameWindows[2].Width = FRAME_THICKNESS;
                        _frameWindows[2].Height = atTopEdge && atBottomEdge ? height - (2 * FRAME_THICKNESS) : 
                                                 atTopEdge ? height - FRAME_THICKNESS :
                                                 atBottomEdge ? height - FRAME_THICKNESS : height;
                    }
                    else
                    {
                        _frameWindows[2].Left = left - FRAME_THICKNESS;
                        _frameWindows[2].Top = top;
                        _frameWindows[2].Width = FRAME_THICKNESS;
                        _frameWindows[2].Height = height;
                    }
                    
                    // Right frame - if at right edge, position inside the window
                    if (atRightEdge)
                    {
                        _frameWindows[3].Left = left + width - FRAME_THICKNESS;
                        _frameWindows[3].Top = atTopEdge ? top + FRAME_THICKNESS : top;
                        _frameWindows[3].Width = FRAME_THICKNESS;
                        _frameWindows[3].Height = atTopEdge && atBottomEdge ? height - (2 * FRAME_THICKNESS) : 
                                                 atTopEdge ? height - FRAME_THICKNESS :
                                                 atBottomEdge ? height - FRAME_THICKNESS : height;
                    }
                    else
                    {
                        _frameWindows[3].Left = left + width;
                        _frameWindows[3].Top = top;
                        _frameWindows[3].Width = FRAME_THICKNESS;
                        _frameWindows[3].Height = height;
                    }

                    // Enhanced Z-order management - keep frames above target window but below other windows
                    // Find the window immediately above the target window
                    IntPtr windowAbove = NativeMethods.GetWindow(_targetHwnd, NativeMethods.GW_HWNDPREV);
                    
                    // Get handles for frame windows
                    var frameHandles = _frameWindows.Select(fw => new WindowInteropHelper(fw).Handle).ToArray();
                    
                    if (windowAbove != IntPtr.Zero)
                    {
                        // Position frames just above the target window but below the next window
                        foreach (var frameHandle in frameHandles)
                        {
                            NativeMethods.SetWindowPos(frameHandle, windowAbove, 0, 0, 0, 0, 
                                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
                        }
                    }
                    else
                    {
                        // Target window is topmost, so put frames just above it
                        foreach (var frameHandle in frameHandles)
                        {
                            NativeMethods.SetWindowPos(frameHandle, _targetHwnd, 0, 0, 0, 0, 
                                NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOACTIVATE);
                        }
                    }
                });

                Thread.Sleep(16);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            StopTracking();
            base.OnClosed(e);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using WF = System.Windows.Forms;
using SW = System.Windows;
using SWI = System.Windows.Input;
using SWC = System.Windows.Controls;
using SWIop = System.Windows.Interop;
using System.Windows.Threading;

namespace TaskbarPlayer
{
    public partial class MainWindow : SW.Window
    {
        WF.NotifyIcon tray;
        bool loop = true, isMoving = false, isResizing = false, isFull = false, isPaused = false;
        bool canMove = true, clickThrough = false, isMenuOpen = false;

        SW.Rect normalBounds;
        SW.Point moveStart, resizeStart;
        SW.Size startSize;

        List<string> playlist = new();
        int currentIndex = -1;
        ListWindow? listWindow;

        // シークUI・最前面維持
        readonly DispatcherTimer posTimer = new DispatcherTimer();
        readonly DispatcherTimer topTimer = new DispatcherTimer();
        TimeSpan mediaDuration = TimeSpan.Zero;
        SWC.Slider? seekSliderWpf;     // 右クリックメニュー内スライダー
        WF.TrackBar? seekBarTray;      // タスクトレイ内スライダー

        public MainWindow()
        {
            InitializeComponent();
            InitTray();

            SourceInitialized += (_, __) =>
            {
                ApplyToolWindowStyle();
                EnsureTopMost();
            };
            Activated += (_, __) => EnsureTopMost();
            Deactivated += (_, __) => EnsureTopMost();
            LocationChanged += (_, __) => EnsureTopMost();
            StateChanged += (_, __) => EnsureTopMost();
            SizeChanged += (_, __) => EnsureTopMost();

            posTimer.Interval = TimeSpan.FromMilliseconds(500);
            posTimer.Tick += (_, __) => UpdateSeekUI();
            posTimer.Start();

            topTimer.Interval = TimeSpan.FromMilliseconds(800);
            topTimer.Tick += (_, __) => { if (!isMenuOpen) EnsureTopMost(); };
            topTimer.Start();

            Player.MediaOpened += (_, __) =>
            {
                mediaDuration = Player.NaturalDuration.HasTimeSpan ? Player.NaturalDuration.TimeSpan : TimeSpan.Zero;
                UpdateSeekUI();
            };
        }

        // ===== ダブルクリック：全画面トグル =====
        private void OnDoubleClick(object sender, SWI.MouseButtonEventArgs e)
        {
            if (!isFull)
            {
                normalBounds = new SW.Rect(Left, Top, Width, Height);
                Left = 0; Top = 0;
                Width = SW.SystemParameters.PrimaryScreenWidth;
                Height = SW.SystemParameters.PrimaryScreenHeight;
                isFull = true;
            }
            else
            {
                Left = normalBounds.Left; Top = normalBounds.Top;
                Width = normalBounds.Width; Height = normalBounds.Height;
                isFull = false;
            }
            EnsureTopMost();
        }

        // ===== 左ドラッグ移動/右下リサイズ =====
        private void OnLeftDown(object sender, SWI.MouseButtonEventArgs e)
        {
            if (!canMove) return;
            const int edge = 12;
            var pos = e.GetPosition(this);
            bool nearCorner = (pos.X >= Width - edge && pos.Y >= Height - edge);

            if (nearCorner)
            {
                isResizing = true;
                resizeStart = pos;
                startSize = new SW.Size(Width, Height);
                SWI.Mouse.Capture(this);
            }
            else
            {
                isMoving = true;
                moveStart = pos;
                SWI.Mouse.Capture(this);
            }
        }

        private void OnLeftUp(object sender, SWI.MouseButtonEventArgs e)
        {
            if (isMoving || isResizing)
            {
                isMoving = false; isResizing = false;
                SWI.Mouse.Capture(null);
            }
            EnsureTopMost();
        }

        private void OnResizeMove(object sender, SWI.MouseEventArgs e)
        {
            if (isFull || !canMove) return;
            const int edge = 12;
            var pos = e.GetPosition(this);
            bool nearCorner = (pos.X >= Width - edge && pos.Y >= Height - edge);
            Cursor = nearCorner ? SWI.Cursors.SizeNWSE : SWI.Cursors.Arrow;

            if (isResizing && e.LeftButton == SWI.MouseButtonState.Pressed)
            {
                var diff = e.GetPosition(this) - resizeStart;
                double newWidth = Math.Max(40, startSize.Width + diff.X);
                double newHeight = newWidth * (startSize.Height / startSize.Width);
                Width = newWidth; Height = newHeight;
            }

            if (isMoving && e.LeftButton == SWI.MouseButtonState.Pressed)
            {
                var posNow = e.GetPosition(this);
                Left += posNow.X - moveStart.X;
                Top += posNow.Y - moveStart.Y;
            }
        }

        // ===== D&D追加 =====
        private void OnFileDrop(object sender, SW.DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(SW.DataFormats.FileDrop)) return;
            string[] files = (string[])e.Data.GetData(SW.DataFormats.FileDrop);
            foreach (var f in files) AddToPlaylist(f);
            if (files.Length > 0) PlayFromIndex(playlist.IndexOf(files[0]));
        }

        // ===== 再生 =====
        void PlayFromIndex(int index)
        {
            if (index < 0 || index >= playlist.Count) return;
            currentIndex = index;
            Player.Source = new Uri(playlist[index]);
            Player.MediaEnded -= OnMediaEnded;
            Player.MediaEnded += OnMediaEnded;
            Player.Play();
            isPaused = false;
            EnsureTopMost();
        }

        void OnMediaEnded(object sender, SW.RoutedEventArgs e)
        {
            if (playlist.Count > 1)
            {
                currentIndex = (currentIndex + 1) % playlist.Count;
                PlayFromIndex(currentIndex);
            }
            else if (loop)
            {
                Player.Position = TimeSpan.Zero;
                Player.Play();
            }
        }

        void TogglePause()
        {
            if (isPaused) { Player.Play(); isPaused = false; }
            else { Player.Pause(); isPaused = true; }
        }

        void AddToPlaylist(string path)
        {
            playlist.Add(path);
            listWindow?.SyncList(playlist);
        }

        // ===== 再生リスト =====
        void OpenListWindow()
        {
            if (listWindow == null)
            {
                listWindow = new ListWindow(playlist, PlayFromIndex);
                listWindow.Owner = this;
                listWindow.Closing += (s, e) => { e.Cancel = true; listWindow.Hide(); };
                listWindow.Show();
            }
            else
            {
                if (!listWindow.IsVisible) listWindow.Show(); else listWindow.Activate();
            }
            EnsureTopMost();
        }

        // ===== 右クリックメニュー（WPF） =====
        private void OnRightClick(object sender, SWI.MouseButtonEventArgs e)
        {
            var cm = BuildWpfMenu();
            isMenuOpen = true;
            cm.Closed += (_, __) => { isMenuOpen = false; EnsureTopMost(); };
            cm.IsOpen = true; // ContextMenu自体は最前面ウインドウなのでプレーヤーより前に出る
        }

        SWC.ContextMenu BuildWpfMenu()
        {
            var cm = new SWC.ContextMenu();

            var open = new SWC.MenuItem { Header = "開く..." };
            open.Click += (_, __) =>
            {
                var dlg = new OpenFileDialog { Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|すべて|*.*" };
                if (dlg.ShowDialog() == true) { AddToPlaylist(dlg.FileName); PlayFromIndex(playlist.IndexOf(dlg.FileName)); }
            };

            var listItem = new SWC.MenuItem { Header = "リストウインドウ" };
            listItem.Click += (_, __) => OpenListWindow();

            var pauseItem = new SWC.MenuItem { Header = "一時停止 / 再開" };
            pauseItem.Click += (_, __) => TogglePause();

            // 再生位置（秒）
            seekSliderWpf = new SWC.Slider { Minimum = 0, Maximum = 0, Value = 0, Width = 160, Margin = new SW.Thickness(6) };
            seekSliderWpf.ValueChanged += (_, __) =>
            {
                if (seekSliderWpf.IsMouseCaptureWithin) Player.Position = TimeSpan.FromSeconds(seekSliderWpf.Value);
            };
            var seekItem = new SWC.MenuItem { Header = "位置" };
            seekItem.Items.Add(seekSliderWpf);
            seekItem.StaysOpenOnClick = true;

            // 音量
            var volSlider = new SWC.Slider { Minimum = 0, Maximum = 1, Value = Player.Volume, Width = 120, Margin = new SW.Thickness(6) };
            volSlider.ValueChanged += (_, __) => Player.Volume = volSlider.Value;
            var volItem = new SWC.MenuItem { Header = "音量" };
            volItem.Items.Add(volSlider);
            volItem.StaysOpenOnClick = true;

            // 透明度
            var opSlider = new SWC.Slider { Minimum = 0.0, Maximum = 1.0, Value = this.Opacity, Width = 120, Margin = new SW.Thickness(6) };
            opSlider.ValueChanged += (_, __) => this.Opacity = opSlider.Value;
            var opItem = new SWC.MenuItem { Header = "透明度" };
            opItem.Items.Add(opSlider);
            opItem.StaysOpenOnClick = true;

            var topItem = new SWC.MenuItem { Header = "最前列に表示" };
            topItem.Click += (_, __) => EnsureTopMost();

            var moveItem = new SWC.MenuItem { Header = "ウインドウの移動許可", IsCheckable = true, IsChecked = canMove };
            moveItem.Checked += (_, __) => { canMove = true; ToggleClickThrough(false); };
            moveItem.Unchecked += (_, __) => { canMove = false; ToggleClickThrough(true); };

            var quit = new SWC.MenuItem { Header = "終了" };
            quit.Click += (_, __) => { tray.Visible = false; tray.Dispose(); Close(); };

            cm.Items.Add(open);
            cm.Items.Add(listItem);
            cm.Items.Add(pauseItem);
            cm.Items.Add(seekItem);
            cm.Items.Add(volItem);
            cm.Items.Add(opItem);
            cm.Items.Add(topItem);
            cm.Items.Add(moveItem);
            cm.Items.Add(new SWC.Separator());
            cm.Items.Add(quit);
            return cm;
        }

        // ===== タスクトレイ =====
        void InitTray()
        {
            var iconUri = new Uri("pack://application:,,,/TaskbarPlayer;component/TaskVideoPlayer.ico");
            var iconStream = SW.Application.GetResourceStream(iconUri)?.Stream;

            tray = new WF.NotifyIcon
            {
                Icon = iconStream != null ? new System.Drawing.Icon(iconStream) : System.Drawing.SystemIcons.Application,
                Visible = true,
                Text = "Taskbar Player"
            };

            tray.MouseClick += (_, e) =>
            {
                if (e.Button == WF.MouseButtons.Right)
                {
                    var cm = BuildTrayMenu();
                    isMenuOpen = true;
                    cm.Closed += (_, __) => { isMenuOpen = false; EnsureTopMost(); };
                    tray.ContextMenuStrip = cm;
                    cm.Show(WF.Cursor.Position);
                }
                else if (e.Button == WF.MouseButtons.Left)
                {
                    OpenListWindow();
                    EnsureTopMost();
                }
            };
        }

        WF.ContextMenuStrip BuildTrayMenu()
        {
            var cm = new WF.ContextMenuStrip();

            var open = new WF.ToolStripMenuItem("開く...");
            open.Click += (_, __) =>
            {
                var dlg = new OpenFileDialog { Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|すべて|*.*" };
                if (dlg.ShowDialog() == true) { AddToPlaylist(dlg.FileName); PlayFromIndex(playlist.IndexOf(dlg.FileName)); }
            };

            var listItem = new WF.ToolStripMenuItem("リストウインドウ");
            listItem.Click += (_, __) => OpenListWindow();

            var pause = new WF.ToolStripMenuItem("一時停止 / 再開");
            pause.Click += (_, __) => TogglePause();

            // 位置（秒）
            seekBarTray = new WF.TrackBar { Minimum = 0, Maximum = 1, TickStyle = WF.TickStyle.None, Value = 0, Width = 160 };
            seekBarTray.Scroll += (_, __) => { Player.Position = TimeSpan.FromSeconds(seekBarTray.Value); };
            var seekHost = new WF.ToolStripControlHost(seekBarTray);
            var seekLabel = new WF.ToolStripMenuItem("位置（秒）") { Enabled = false };

            // 音量・透明度
            var volTrack = new WF.TrackBar { Minimum = 0, Maximum = 100, TickStyle = WF.TickStyle.None, Value = (int)(Player.Volume * 100), Width = 120 };
            volTrack.ValueChanged += (_, __) => Player.Volume = volTrack.Value / 100.0;
            var volHost = new WF.ToolStripControlHost(volTrack);
            var volLabel = new WF.ToolStripMenuItem("音量") { Enabled = false };

            var opTrack = new WF.TrackBar { Minimum = 0, Maximum = 100, TickStyle = WF.TickStyle.None, Value = (int)(this.Opacity * 100), Width = 120 };
            opTrack.ValueChanged += (_, __) => this.Opacity = opTrack.Value / 100.0;
            var opHost = new WF.ToolStripControlHost(opTrack);
            var opLabel = new WF.ToolStripMenuItem("透明度") { Enabled = false };

            var topItem = new WF.ToolStripMenuItem("最前列に表示");
            topItem.Click += (_, __) => EnsureTopMost();

            var moveItem = new WF.ToolStripMenuItem("ウインドウの移動許可") { Checked = canMove, CheckOnClick = true };
            moveItem.CheckedChanged += (_, __) => { canMove = moveItem.Checked; ToggleClickThrough(!canMove); };

            var quit = new WF.ToolStripMenuItem("終了");
            quit.Click += (_, __) => { tray.Visible = false; tray.Dispose(); Close(); };

            cm.Items.Add(open);
            cm.Items.Add(listItem);
            cm.Items.Add(pause);
            cm.Items.Add(new WF.ToolStripSeparator());
            cm.Items.Add(seekLabel); cm.Items.Add(seekHost);
            cm.Items.Add(volLabel); cm.Items.Add(volHost);
            cm.Items.Add(opLabel); cm.Items.Add(opHost);
            cm.Items.Add(new WF.ToolStripSeparator());
            cm.Items.Add(topItem);
            cm.Items.Add(moveItem);
            cm.Items.Add(new WF.ToolStripSeparator());
            cm.Items.Add(quit);
            return cm;
        }

        // ===== シークUI同期 =====
        void UpdateSeekUI()
        {
            double sec = Player.Position.TotalSeconds;
            double dur = mediaDuration.TotalSeconds > 0 ? mediaDuration.TotalSeconds : 0;

            if (seekSliderWpf != null)
            {
                if (dur > 0 && Math.Abs(seekSliderWpf.Maximum - dur) > 0.5)
                    seekSliderWpf.Maximum = dur;
                if (!seekSliderWpf.IsMouseCaptureWithin)
                    seekSliderWpf.Value = Math.Max(0, Math.Min(sec, seekSliderWpf.Maximum));
            }
            if (seekBarTray != null)
            {
                int max = (int)Math.Max(1, Math.Round(dur));
                if (seekBarTray.Maximum != max) seekBarTray.Maximum = max;
                if (!seekBarTray.Capture)
                {
                    int v = (int)Math.Max(0, Math.Min(sec, seekBarTray.Maximum));
                    if (v >= seekBarTray.Minimum && v <= seekBarTray.Maximum) seekBarTray.Value = v;
                }
            }
        }

        // ===== クリック透過 =====
        void ToggleClickThrough(bool enable)
        {
            var hwnd = new SWIop.WindowInteropHelper(this).Handle;
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            if (enable) { _ = SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TRANSPARENT); }
            else { _ = SetWindowLong(hwnd, GWL_EXSTYLE, exStyle & ~WS_EX_TRANSPARENT); }
            clickThrough = enable;
            EnsureTopMost();
        }

        // ===== ツールウインドウ + TopMost =====
        void ApplyToolWindowStyle()
        {
            var hwnd = new SWIop.WindowInteropHelper(this).Handle;
            int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
            _ = SetWindowLong(hwnd, GWL_EXSTYLE, ex | WS_EX_TOOLWINDOW | WS_EX_TOPMOST);
        }

        // ===== タスクバーより上へ維持 =====
        void EnsureTopMost()
        {
            if (isMenuOpen) return; // メニュー最前面を妨げない.
            try
            {
                var hwnd = new SWIop.WindowInteropHelper(this).Handle;
                if (hwnd == IntPtr.Zero) return;
                _ = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0,
                    SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER | SWP_NOSENDCHANGING | SWP_SHOWWINDOW);
            }
            catch { }
            Topmost = true;
        }

        // Win32
        const int GWL_EXSTYLE = -20;
        const int WS_EX_TOOLWINDOW = 0x00000080;
        const int WS_EX_TOPMOST = 0x00000008;
        const int WS_EX_TRANSPARENT = 0x00000020;

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        const uint SWP_NOSIZE = 0x0001;
        const uint SWP_NOMOVE = 0x0002;
        const uint SWP_NOACTIVATE = 0x0010;
        const uint SWP_SHOWWINDOW = 0x0040;
        const uint SWP_NOOWNERZORDER = 0x0200;
        const uint SWP_NOSENDCHANGING = 0x0400;

        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);
    }

    // ===== 再生リスト =====
    public class ListWindow : SW.Window
    {
        SWC.ListBox list = new();
        List<string> playlist;
        Action<int> playAction;

        public ListWindow(List<string> sharedList, Action<int> playAction)
        {
            this.playlist = sharedList;
            this.playAction = playAction;

            Title = "再生リスト";
            Width = 300; Height = 260;
            WindowStartupLocation = SW.WindowStartupLocation.CenterOwner;
            WindowStyle = SW.WindowStyle.ToolWindow;
            Topmost = true;
            AllowDrop = true;
            Drop += OnDropFiles;

            var grid = new SWC.Grid();
            var label = new SWC.TextBlock
            {
                Text = "全曲リピート再生します。",
                Margin = new SW.Thickness(8),
                HorizontalAlignment = SW.HorizontalAlignment.Center
            };

            list.Margin = new SW.Thickness(4);
            list.Height = 150;
            list.MouseDoubleClick += (_, __) => { if (list.SelectedIndex >= 0) playAction(list.SelectedIndex); };

            var upBtn = new SWC.Button { Content = "▲", Width = 40, Margin = new SW.Thickness(4) };
            var downBtn = new SWC.Button { Content = "▼", Width = 40, Margin = new SW.Thickness(4) };
            upBtn.Click += (_, __) => MoveItem(-1);
            downBtn.Click += (_, __) => MoveItem(1);

            var addBtn = new SWC.Button { Content = "追加", Margin = new SW.Thickness(4) };
            var delBtn = new SWC.Button { Content = "削除", Margin = new SW.Thickness(4) };
            addBtn.Click += (_, __) =>
            {
                var dlg = new OpenFileDialog { Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|すべて|*.*", Multiselect = true };
                if (dlg.ShowDialog() == true) foreach (var f in dlg.FileNames) AddFile(f);
            };
            delBtn.Click += (_, __) => { if (list.SelectedItem is string file) { playlist.Remove(file); Refresh(); } };

            var panel = new SWC.StackPanel { Orientation = SWC.Orientation.Horizontal, HorizontalAlignment = SW.HorizontalAlignment.Center };
            panel.Children.Add(addBtn); panel.Children.Add(delBtn);
            panel.Children.Add(upBtn); panel.Children.Add(downBtn);

            grid.RowDefinitions.Add(new SWC.RowDefinition { Height = SW.GridLength.Auto });
            grid.RowDefinitions.Add(new SWC.RowDefinition());
            grid.RowDefinitions.Add(new SWC.RowDefinition { Height = SW.GridLength.Auto });

            grid.Children.Add(label); SWC.Grid.SetRow(label, 0);
            grid.Children.Add(list); SWC.Grid.SetRow(list, 1);
            grid.Children.Add(panel); SWC.Grid.SetRow(panel, 2);
            Content = grid;

            Closing += (s, e) => { e.Cancel = true; Hide(); };
            Refresh();
        }

        void OnDropFiles(object sender, SW.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(SW.DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(SW.DataFormats.FileDrop);
                foreach (var f in files) AddFile(f);
            }
        }

        void AddFile(string path) { playlist.Add(path); Refresh(); }

        void MoveItem(int direction)
        {
            int i = list.SelectedIndex; if (i < 0) return;
            int ni = i + direction; if (ni < 0 || ni >= playlist.Count) return;
            var item = playlist[i];
            playlist.RemoveAt(i);
            playlist.Insert(ni, item);
            Refresh();
            list.SelectedIndex = ni;
        }

        public void SyncList(List<string> src) { playlist = src; Refresh(); }

        void Refresh()
        {
            list.Items.Clear();
            foreach (var f in playlist) list.Items.Add(f);
        }
    }
}

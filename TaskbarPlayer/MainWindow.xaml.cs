using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Threading;
using SW = System.Windows;
using SWC = System.Windows.Controls;
using SWI = System.Windows.Input;
using SWIop = System.Windows.Interop;
using WF = System.Windows.Forms;

namespace TubePlayer
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
        ListWindow listWindow;

        // 失敗済みエントリ（自動スキップ用）
        readonly HashSet<string> failedEntries = new();

        // シークUI・最前面維持
        readonly DispatcherTimer posTimer = new DispatcherTimer();
        readonly DispatcherTimer topTimer = new DispatcherTimer();
        TimeSpan mediaDuration = TimeSpan.Zero;
        SWC.Slider seekSliderWpf;     // 右クリックメニュー内スライダー
        WF.TrackBar seekBarTray;      // タスクトレイ内スライダー

        // ステータスメッセージの点滅用
        readonly DispatcherTimer statusTimer = new DispatcherTimer();

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
            topTimer.Tick += (_, __) => { if (!IsPopupOrListOpen) EnsureTopMost(); };
            topTimer.Start();

            // メディア読み込み完了時：ステータスを消す
            Player.MediaOpened += (_, __) =>
            {
                mediaDuration = Player.NaturalDuration.HasTimeSpan ? Player.NaturalDuration.TimeSpan : TimeSpan.Zero;
                HideStatus();
                UpdateSeekUI();
            };

            // メディア読み込み失敗時
            Player.MediaFailed += async (_, e) =>
            {
                string failed = (currentIndex >= 0 && currentIndex < playlist.Count)
                    ? playlist[currentIndex]
                    : "";

                string shortInfo = string.IsNullOrEmpty(failed) ? "" : "\n" + Shorten(failed, 60);

                ShowStatus("取得に失敗しました" + shortInfo);
                SW.MessageBox.Show(
                    "メディアの読み込みに失敗しました。\n" + (e.ErrorException?.Message ?? ""),
                    "再生エラー");

                if (!string.IsNullOrEmpty(failed) && currentIndex >= 0 && currentIndex < playlist.Count)
                {
                    failedEntries.Add(failed);
                    playlist[currentIndex] = "[取得失敗] " + failed;
                    listWindow?.SyncList(playlist);
                    await TryPlayNextAfterFailureAsync(failed);
                }
            };

            // ステータスメッセージ点滅（1秒間隔）
            statusTimer.Interval = TimeSpan.FromSeconds(1);
            statusTimer.Tick += (_, __) =>
            {
                if (StatusOverlay.Visibility != SW.Visibility.Visible) return;
                StatusOverlay.Opacity = (StatusOverlay.Opacity > 0.5) ? 0.3 : 1.0;
            };
        }

        // 起動時に本体ウインドウへフォーカスを当てて、ペースト（Ctrl+V）を確実に拾う
        private void OnLoaded(object sender, SW.RoutedEventArgs e)
        {
            Focus();
        }

        // ★ 画面上メッセージ表示／非表示
        void ShowStatus(string message)
        {
            StatusLabel.Text = message;
            StatusOverlay.Visibility = SW.Visibility.Visible;
            StatusOverlay.Opacity = 1.0;
            statusTimer.Start();
        }

        void HideStatus()
        {
            statusTimer.Stop();
            StatusOverlay.Visibility = SW.Visibility.Collapsed;
        }

        static string Shorten(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= max ? s : s.Substring(0, max) + "...";
        }

        // ===== Ctrl+V 受付（本体） =====
        private async void OnPreviewKeyDown(object sender, SWI.KeyEventArgs e)
        {
            if (e.Key == SWI.Key.V && (SWI.Keyboard.Modifiers & SWI.ModifierKeys.Control) == SWI.ModifierKeys.Control)
            {
                if (WF.Clipboard.ContainsText())
                {
                    var text = WF.Clipboard.GetText();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        await AddFromTextAsync(text);
                        e.Handled = true;
                    }
                }
            }
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

        // ===== D&D（本体：テキストURL / .url / .txt / 動画ファイル） =====
        private async void OnFileDrop(object sender, SW.DragEventArgs e)
        {
            // 1) テキスト（UnicodeText / Text 両方を見る）
            string text = GetDropText(e.Data);
            if (!string.IsNullOrWhiteSpace(text))
            {
                await AddFromTextAsync(text);
                return;
            }

            // 2) ファイル群
            if (e.Data.GetDataPresent(SW.DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(SW.DataFormats.FileDrop);
                await AddFromFilesAsync(files);
            }
        }

        // ====== 共有：テキストから追加（本体・リスト共通） ======
        private async Task AddFromTextAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            bool isIdle = (currentIndex < 0);   // 何も再生していないときだけオーバーレイと自動再生

            if (isIdle)
                ShowStatus("動画読み込み中…");

            int before = playlist.Count;
            await ProcessTextLinesAsync(text);
            int added = playlist.Count - before;

            if (isIdle)
            {
                if (added > 0)
                {
                    await PlayFromIndexAsync(before);
                }
                else
                {
                    ShowStatus("取得に失敗しました");
                }
            }
            // 再生中（!isIdle）の場合は、リストにだけ追加して再生中の動画はそのまま
        }

        // ====== 共有：ファイル群から追加（本体・リスト共通） ======
        private async Task AddFromFilesAsync(string[] files)
        {
            if (files == null || files.Length == 0) return;

            bool isIdle = (currentIndex < 0);   // 何も再生していないときだけオーバーレイと自動再生

            if (isIdle)
                ShowStatus("動画読み込み中…");

            int before = playlist.Count;
            bool anyAdded = false;

            // 1) .url 展開
            foreach (var f in files)
            {
                if (IsInternetShortcut(f) && TryExtractUrlFromInternetShortcut(f, out var url))
                {
                    AddToPlaylist(url);
                    anyAdded = true;
                }
            }

            // 2) .txt は行として読み込む
            foreach (var f in files.Where(x =>
                         string.Equals(Path.GetExtension(x), ".txt", StringComparison.OrdinalIgnoreCase)))
            {
                string text = SafeReadAllText(f);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    await ProcessTextLinesAsync(text);
                    anyAdded = true;
                }
            }

            // 3) 動画ファイル
            foreach (var f in files)
            {
                if (IsInternetShortcut(f)) continue;
                if (string.Equals(Path.GetExtension(f), ".txt", StringComparison.OrdinalIgnoreCase)) continue;

                if (IsLikelyVideo(f))
                {
                    AddToPlaylist(f);
                    anyAdded = true;
                }
            }

            int added = playlist.Count - before;

            if (isIdle)
            {
                if ((anyAdded || added > 0))
                {
                    await PlayFromIndexAsync(before);
                }
                else
                {
                    ShowStatus("取得に失敗しました");
                }
            }
            // 再生中（!isIdle）の場合は、リストだけ増えて再生はそのまま
        }

        // 行単位でURL/パスを処理
        private async Task<int> ProcessTextLinesAsync(string multi)
        {
            if (string.IsNullOrWhiteSpace(multi)) return 0;
            int before = playlist.Count;

            var lines = multi.Split(new[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var raw in lines)
            {
                foreach (var token in raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var s = token.Trim().Trim('"');
                    if (string.IsNullOrEmpty(s)) continue;

                    // URL
                    if (IsWebUrl(s))
                    {
                        AddToPlaylist(s);
                        continue;
                    }

                    // フルパス or 相対 -> 絶対に
                    string candidate = s;
                    if (!Path.IsPathRooted(candidate))
                    {
                        try { candidate = Path.GetFullPath(candidate); } catch { }
                    }
                    if (File.Exists(candidate))
                    {
                        // .url は展開
                        if (IsInternetShortcut(candidate) &&
                            TryExtractUrlFromInternetShortcut(candidate, out var u))
                        {
                            AddToPlaylist(u);
                            continue;
                        }
                        // .txt は再帰的に読む
                        if (string.Equals(Path.GetExtension(candidate), ".txt", StringComparison.OrdinalIgnoreCase))
                        {
                            string text = SafeReadAllText(candidate);
                            await ProcessTextLinesAsync(text);
                            continue;
                        }
                        // 動画
                        if (IsLikelyVideo(candidate))
                        {
                            AddToPlaylist(candidate);
                            continue;
                        }
                    }
                }
            }
            listWindow?.SyncList(playlist);
            return playlist.Count - before;
        }

        // ===== 再生 =====
        private async Task PlayFromIndexAsync(int index)
        {
            if (playlist.Count == 0)
            {
                Player.Stop();
                Player.Source = null;
                currentIndex = -1;
                ShowStatus("取得に失敗しました");
                return;
            }

            if (index < 0 || index >= playlist.Count) return;

            currentIndex = index;
            string entry = playlist[index];

            // 失敗マーク済みなら次へ
            if (!string.IsNullOrWhiteSpace(entry) && entry.StartsWith("[取得失敗]"))
            {
                await TryPlayNextAfterFailureAsync(entry);
                return;
            }

            ShowStatus("動画読み込み中…\n" + Shorten(entry, 60));

            string src = entry;

            try
            {
                Player.Source = new Uri(src);
            }
            catch
            {
                Player.Source = null;
                failedEntries.Add(entry);
                playlist[index] = "[取得失敗] " + entry;
                listWindow?.SyncList(playlist);
                ShowStatus("取得に失敗しました\n" + Shorten(entry, 60));
                await TryPlayNextAfterFailureAsync(entry);
                return;
            }

            Player.MediaEnded -= OnMediaEnded;
            Player.MediaEnded += OnMediaEnded;
            Player.Play();
            isPaused = false;
            EnsureTopMost();
        }

        // 失敗時に次のアイテムへ自動的に進む
        private async Task TryPlayNextAfterFailureAsync(string failedEntry)
        {
            failedEntries.Add(failedEntry);

            if (playlist.Count == 0)
            {
                ShowStatus("取得に失敗しました");
                Player.Stop();
                Player.Source = null;
                currentIndex = -1;
                isPaused = false;
                return;
            }

            int start = currentIndex;
            if (start < 0) start = 0;

            for (int offset = 1; offset <= playlist.Count; offset++)
            {
                int idx = (start + offset) % playlist.Count;
                if (idx < 0 || idx >= playlist.Count) continue;

                string item = playlist[idx];
                if (string.IsNullOrWhiteSpace(item)) continue;
                if (item.StartsWith("[取得失敗]")) continue;
                if (failedEntries.Contains(item)) continue;

                await PlayFromIndexAsync(idx);
                return;
            }

            // 全部ダメ
            ShowStatus("取得に失敗しました");
            Player.Stop();
            Player.Source = null;
            currentIndex = -1;
            isPaused = false;
        }

        private void OnMediaEnded(object sender, SW.RoutedEventArgs e)
        {
            _ = OnMediaEndedAsync();
        }

        private async Task OnMediaEndedAsync()
        {
            if (playlist.Count > 1)
            {
                int next = (currentIndex + 1 + playlist.Count) % playlist.Count;
                await PlayFromIndexAsync(next);
            }
            else if (loop && playlist.Count == 1 &&
                     currentIndex >= 0 && currentIndex < playlist.Count &&
                     !playlist[currentIndex].StartsWith("[取得失敗]"))
            {
                // 単一アイテムのループ
                Player.Position = TimeSpan.Zero;
                Player.Play();
            }
        }

        private void TogglePause()
        {
            if (isPaused) { Player.Play(); isPaused = false; }
            else { Player.Pause(); isPaused = true; }
        }

        private void AddToPlaylist(string pathOrUrl)
        {
            if (!playlist.Contains(pathOrUrl))
            {
                playlist.Add(pathOrUrl);
                listWindow?.SyncList(playlist);
            }
        }

        // ★ リスト全削除（本体側の統一処理）
        private void ClearList()
        {
            try { Player.Stop(); } catch { }
            Player.Source = null;
            HideStatus();
            playlist.Clear();
            currentIndex = -1;
            mediaDuration = TimeSpan.Zero;
            failedEntries.Clear();

            // シークUIリセット
            if (seekSliderWpf != null)
            {
                seekSliderWpf.Maximum = 0;
                seekSliderWpf.Value = 0;
            }
            if (seekBarTray != null)
            {
                seekBarTray.Maximum = 1;
                seekBarTray.Value = 0;
            }

            listWindow?.SyncList(playlist);
        }

        // ===== 再生リスト =====
        private void OpenListWindow()
        {
            if (listWindow == null)
            {
                listWindow = new ListWindow(
                    playlist,
                    index => _ = PlayFromIndexAsync(index),
                    SaveListToTxt,
                    LoadListFromTxt,
                    addTextAsync: AddFromTextAsync,
                    addFilesAsync: AddFromFilesAsync,
                    clearAction: ClearList
                );

                listWindow.Owner = this;
                listWindow.Topmost = true;
                listWindow.ShowInTaskbar = false;
                listWindow.Closing += (s, e) => { e.Cancel = true; listWindow.Hide(); };
                listWindow.IsVisibleChanged += (_, __) => EnsureTopMost();
                listWindow.Show();
                listWindow.Activate();
            }
            else
            {
                if (!listWindow.IsVisible) listWindow.Show();
                listWindow.Topmost = true;
                listWindow.Activate();
            }
            EnsureTopMost();
        }

        // ===== 右クリックメニュー（WPF） =====
        private void OnRightClick(object sender, SWI.MouseButtonEventArgs e)
        {
            var cm = BuildWpfMenu();
            isMenuOpen = true;
            cm.Closed += (_, __) => { isMenuOpen = false; EnsureTopMost(); };
            cm.IsOpen = true;
        }

        private SWC.ContextMenu BuildWpfMenu()
        {
            var cm = new SWC.ContextMenu();

            var open = new SWC.MenuItem { Header = "開く..." };
            open.Click += async (_, __) =>
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|テキスト|*.txt|すべて|*.*",
                    Multiselect = true
                };
                if (dlg.ShowDialog() == true)
                {
                    await AddFromFilesAsync(dlg.FileNames);
                }
            };

            var listItem = new SWC.MenuItem { Header = "リストウインドウ" };
            listItem.Click += (_, __) => OpenListWindow();

            var saveList = new SWC.MenuItem { Header = "リストを保存..." };
            saveList.Click += (_, __) => SaveListToTxt();

            var loadList = new SWC.MenuItem { Header = "リストを読み込み..." };
            loadList.Click += (_, __) => LoadListFromTxt();

            var clearList = new SWC.MenuItem { Header = "リストをクリア" };
            clearList.Click += (_, __) => ClearList();

            var pauseItem = new SWC.MenuItem { Header = "一時停止 / 再開" };
            pauseItem.Click += (_, __) => TogglePause();

            // 再生位置（秒）
            seekSliderWpf = new SWC.Slider
            {
                Minimum = 0,
                Maximum = 0,
                Value = 0,
                Width = 160,
                Margin = new SW.Thickness(6)
            };
            seekSliderWpf.ValueChanged += (_, __) =>
            {
                if (seekSliderWpf.IsMouseCaptureWithin)
                    Player.Position = TimeSpan.FromSeconds(seekSliderWpf.Value);
            };
            var seekItem = new SWC.MenuItem { Header = "位置" };
            seekItem.Items.Add(seekSliderWpf);
            seekItem.StaysOpenOnClick = true;

            // 音量
            var volSlider = new SWC.Slider
            {
                Minimum = 0,
                Maximum = 1,
                Value = Player.Volume,
                Width = 120,
                Margin = new SW.Thickness(6)
            };
            volSlider.ValueChanged += (_, __) => Player.Volume = volSlider.Value;
            var volItem = new SWC.MenuItem { Header = "音量" };
            volItem.Items.Add(volSlider);
            volItem.StaysOpenOnClick = true;

            // 透明度
            var opSlider = new SWC.Slider
            {
                Minimum = 0.0,
                Maximum = 1.0,
                Value = this.Opacity,
                Width = 120,
                Margin = new SW.Thickness(6)
            };
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
            cm.Items.Add(saveList);
            cm.Items.Add(loadList);
            cm.Items.Add(clearList);
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
        private void InitTray()
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

        private WF.ContextMenuStrip BuildTrayMenu()
        {
            var cm = new WF.ContextMenuStrip();

            var open = new WF.ToolStripMenuItem("開く...");
            open.Click += async (_, __) =>
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|テキスト|*.txt|すべて|*.*",
                    Multiselect = true
                };
                if (dlg.ShowDialog() == true)
                {
                    await AddFromFilesAsync(dlg.FileNames);
                }
            };

            var listItem = new WF.ToolStripMenuItem("リストウインドウ");
            listItem.Click += (_, __) => OpenListWindow();

            var saveList = new WF.ToolStripMenuItem("リストを保存...");
            saveList.Click += (_, __) => SaveListToTxt();

            var loadList = new WF.ToolStripMenuItem("リストを読み込み...");
            loadList.Click += (_, __) => LoadListFromTxt();

            var clearList = new WF.ToolStripMenuItem("リストをクリア");
            clearList.Click += (_, __) => ClearList();

            var pause = new WF.ToolStripMenuItem("一時停止 / 再開");
            pause.Click += (_, __) => TogglePause();

            // 位置（秒）
            seekBarTray = new WF.TrackBar
            {
                Minimum = 0,
                Maximum = 1,
                TickStyle = WF.TickStyle.None,
                Value = 0,
                Width = 160
            };
            seekBarTray.Scroll += (_, __) =>
            {
                Player.Position = TimeSpan.FromSeconds(seekBarTray.Value);
            };
            var seekHost = new WF.ToolStripControlHost(seekBarTray);
            var seekLabel = new WF.ToolStripMenuItem("位置（秒）") { Enabled = false };

            // 音量・透明度
            var volTrack = new WF.TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                TickStyle = WF.TickStyle.None,
                Value = (int)(Player.Volume * 100),
                Width = 120
            };
            volTrack.ValueChanged += (_, __) => Player.Volume = volTrack.Value / 100.0;
            var volHost = new WF.ToolStripControlHost(volTrack);
            var volLabel = new WF.ToolStripMenuItem("音量") { Enabled = false };

            var opTrack = new WF.TrackBar
            {
                Minimum = 0,
                Maximum = 100,
                TickStyle = WF.TickStyle.None,
                Value = (int)(this.Opacity * 100),
                Width = 120
            };
            opTrack.ValueChanged += (_, __) => this.Opacity = opTrack.Value / 100.0;
            var opHost = new WF.ToolStripControlHost(opTrack);
            var opLabel = new WF.ToolStripMenuItem("透明度") { Enabled = false };

            var topItem = new WF.ToolStripMenuItem("最前列に表示");
            topItem.Click += (_, __) => EnsureTopMost();

            var moveItem = new WF.ToolStripMenuItem("ウインドウの移動許可") { Checked = canMove, CheckOnClick = true };
            moveItem.CheckedChanged += (_, __) =>
            {
                canMove = moveItem.Checked;
                ToggleClickThrough(!canMove);
            };

            var quit = new WF.ToolStripMenuItem("終了");
            quit.Click += (_, __) => { tray.Visible = false; tray.Dispose(); Close(); };

            cm.Items.Add(open);
            cm.Items.Add(listItem);
            cm.Items.Add(saveList);
            cm.Items.Add(loadList);
            cm.Items.Add(clearList);
            cm.Items.Add(new WF.ToolStripSeparator());
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
        private void UpdateSeekUI()
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
        private void ToggleClickThrough(bool enable)
        {
            var hwnd = new SWIop.WindowInteropHelper(this).Handle;
            int exStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            if (enable) { _ = SetWindowLong(hwnd, GWL_EXSTYLE, exStyle | WS_EX_TRANSPARENT); }
            else { _ = SetWindowLong(hwnd, GWL_EXSTYLE, exStyle & ~WS_EX_TRANSPARENT); }
            clickThrough = enable;
            EnsureTopMost();
        }

        // ===== ツールウインドウ + TopMost =====
        private void ApplyToolWindowStyle()
        {
            var hwnd = new SWIop.WindowInteropHelper(this).Handle;
            int ex = GetWindowLong(hwnd, GWL_EXSTYLE);
            _ = SetWindowLong(hwnd, GWL_EXSTYLE, ex | WS_EX_TOOLWINDOW | WS_EX_TOPMOST);
        }

        // ===== タスクバーより上へ維持 =====
        private void EnsureTopMost()
        {
            if (IsPopupOrListOpen) return; // メニューやリストが最前面であることを妨げない
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

        private bool IsPopupOrListOpen => isMenuOpen || (listWindow != null && listWindow.IsVisible);

        // ===== ヘルパ =====
        private static bool IsLikelyVideo(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return new[] { ".mp4", ".m4v", ".mov", ".avi", ".wmv", ".mkv", ".webm" }.Contains(ext);
        }

        private static bool IsInternetShortcut(string path)
        {
            string e = Path.GetExtension(path).ToLowerInvariant();
            return e == ".url" || e == ".website";
        }

        private static string SafeReadAllText(string path)
        {
            using var fs = File.OpenRead(path);
            using var sr = new StreamReader(fs, Encoding.UTF8, true);
            return sr.ReadToEnd();
        }

        private static bool IsWebUrl(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (!Uri.TryCreate(s.Trim(), UriKind.Absolute, out var uri)) return false;
            return uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps;
        }

        private static bool TryExtractUrlFromInternetShortcut(string filePath, out string url)
        {
            url = "";
            try
            {
                using var fs = File.OpenRead(filePath);
                using var sr = new StreamReader(fs, Encoding.UTF8, true);
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                    {
                        url = line.Substring(4).Trim();
                        return IsWebUrl(url);
                    }
                }
            }
            catch { }
            return false;
        }

        // DataObject からテキストを取り出す（UnicodeText / Text 両対応）
        private static string GetDropText(SW.IDataObject data)
        {
            if (data.GetDataPresent(SW.DataFormats.UnicodeText))
                return data.GetData(SW.DataFormats.UnicodeText) as string;
            if (data.GetDataPresent(SW.DataFormats.Text))
                return data.GetData(SW.DataFormats.Text) as string;
            return null;
        }

        // ===== 保存/読み込み =====
        private void SaveListToTxt()
        {
            var dlg = new SaveFileDialog
            {
                Filter = "テキスト|*.txt",
                FileName = "TaskbarPlayerList.txt",
                AddExtension = true
            };
            if (dlg.ShowDialog() == true)
            {
                File.WriteAllLines(dlg.FileName, playlist, new UTF8Encoding(false));
            }
        }

        private void LoadListFromTxt()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "テキスト|*.txt",
                Multiselect = false
            };
            if (dlg.ShowDialog() == true)
            {
                string text = SafeReadAllText(dlg.FileName);
                _ = AddFromTextAsync(text);
            }
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
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwLong);
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
        Action saveAction;
        Action loadAction;

        // 本体へ委譲するためのコールバック
        Func<string, Task> addTextAsync;
        Func<string[], Task> addFilesAsync;

        // クリア操作
        Action clearAction;

        public ListWindow(
            List<string> sharedList,
            Action<int> playAction,
            Action saveAction,
            Action loadAction,
            Func<string, Task> addTextAsync,
            Func<string[], Task> addFilesAsync,
            Action clearAction
        )
        {
            this.playlist = sharedList;
            this.playAction = playAction;
            this.saveAction = saveAction;
            this.loadAction = loadAction;
            this.addTextAsync = addTextAsync;
            this.addFilesAsync = addFilesAsync;
            this.clearAction = clearAction;

            Title = "再生リスト";
            Width = 360; Height = 300;
            WindowStartupLocation = SW.WindowStartupLocation.CenterOwner;
            WindowStyle = SW.WindowStyle.ToolWindow;
            Topmost = true;
            ShowInTaskbar = false;
            AllowDrop = true;
            Drop += OnDropFiles;
            PreviewKeyDown += OnPreviewKeyDown; // Ctrl+V

            var grid = new SWC.Grid();
            var label = new SWC.TextBlock
            {
                Text = "最前面で表示。D&Dや貼り付けで追加可能。",
                Margin = new SW.Thickness(8),
                HorizontalAlignment = SW.HorizontalAlignment.Center
            };

            list.Margin = new SW.Thickness(4);
            list.Height = 170;
            list.MouseDoubleClick += (_, __) =>
            {
                if (list.SelectedIndex >= 0) playAction(list.SelectedIndex);
            };

            var upBtn = new SWC.Button { Content = "▲", Width = 40, Margin = new SW.Thickness(4) };
            var downBtn = new SWC.Button { Content = "▼", Width = 40, Margin = new SW.Thickness(4) };
            var addBtn = new SWC.Button { Content = "追加", Margin = new SW.Thickness(4) };
            var delBtn = new SWC.Button { Content = "削除", Margin = new SW.Thickness(4) };
            var saveBtn = new SWC.Button { Content = "保存", Margin = new SW.Thickness(4) };
            var loadBtn = new SWC.Button { Content = "読込", Margin = new SW.Thickness(4) };
            var clrBtn = new SWC.Button { Content = "リストクリア", Margin = new SW.Thickness(4) };

            upBtn.Click += (_, __) => MoveItem(-1);
            downBtn.Click += (_, __) => MoveItem(1);
            addBtn.Click += async (_, __) =>
            {
                var dlg = new OpenFileDialog
                {
                    Filter = "動画|*.mp4;*.m4v;*.mov;*.avi;*.wmv|テキスト|*.txt|すべて|*.*",
                    Multiselect = true
                };
                if (dlg.ShowDialog() == true)
                {
                    await addFilesAsync(dlg.FileNames);
                    Refresh();
                }
            };
            delBtn.Click += (_, __) =>
            {
                if (list.SelectedItem is string file)
                {
                    playlist.Remove(file);
                    Refresh();
                }
            };
            saveBtn.Click += (_, __) => saveAction();
            loadBtn.Click += (_, __) => loadAction();
            clrBtn.Click += (_, __) => { clearAction(); Refresh(); };

            var panel = new SWC.StackPanel
            {
                Orientation = SWC.Orientation.Horizontal,
                HorizontalAlignment = SW.HorizontalAlignment.Center
            };
            panel.Children.Add(addBtn); panel.Children.Add(delBtn);
            panel.Children.Add(upBtn); panel.Children.Add(downBtn);
            panel.Children.Add(saveBtn); panel.Children.Add(loadBtn);
            panel.Children.Add(clrBtn);

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

        // Ctrl+V でURL/パスを追加
        private async void OnPreviewKeyDown(object sender, SWI.KeyEventArgs e)
        {
            if (e.Key == SWI.Key.V &&
                (SWI.Keyboard.Modifiers & SWI.ModifierKeys.Control) == SWI.ModifierKeys.Control)
            {
                if (WF.Clipboard.ContainsText())
                {
                    var text = WF.Clipboard.GetText();
                    await addTextAsync(text);
                    Refresh();
                    e.Handled = true;
                }
            }
        }

        // D&Dでテキスト/ファイル追加
        private async void OnDropFiles(object sender, SW.DragEventArgs e)
        {
            // 1) テキスト（UnicodeText / Text 両対応）
            string text = GetDropText(e.Data);
            if (!string.IsNullOrWhiteSpace(text))
            {
                await addTextAsync(text);
                Refresh();
                return;
            }

            // 2) ファイル
            if (e.Data.GetDataPresent(SW.DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(SW.DataFormats.FileDrop);
                await addFilesAsync(files);
                Refresh();
            }
        }

        // DataObject からテキストを取り出す（UnicodeText / Text 両対応）
        private static string GetDropText(SW.IDataObject data)
        {
            if (data.GetDataPresent(SW.DataFormats.UnicodeText))
                return data.GetData(SW.DataFormats.UnicodeText) as string;
            if (data.GetDataPresent(SW.DataFormats.Text))
                return data.GetData(SW.DataFormats.Text) as string;
            return null;
        }

        private void MoveItem(int direction)
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

        private void Refresh()
        {
            list.Items.Clear();
            foreach (var f in playlist) list.Items.Add(f);
        }
    }
}

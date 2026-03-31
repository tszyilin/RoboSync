using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

// ─── Entry point ────────────────────────────────────────────────────────────
static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new MainForm());
    }
}

// ─── Colour palette (Earthy Forest Green) ───────────────────────────────────
static class C
{
    public static readonly Color Bg      = ColorTranslator.FromHtml("#131a0e"); // deep forest dark
    public static readonly Color Surface = ColorTranslator.FromHtml("#1d2a16"); // dark moss green
    public static readonly Color Overlay = ColorTranslator.FromHtml("#2c3e22"); // mid forest green
    public static readonly Color Text    = ColorTranslator.FromHtml("#d4e4bc"); // light sage cream
    public static readonly Color Muted   = ColorTranslator.FromHtml("#6e8858"); // dusty sage
    public static readonly Color Accent  = ColorTranslator.FromHtml("#a0c040"); // olive / lime accent
    public static readonly Color Green   = ColorTranslator.FromHtml("#6ab84e"); // bright forest green
    public static readonly Color Yellow  = ColorTranslator.FromHtml("#c4aa2c"); // golden olive
    public static readonly Color Red     = ColorTranslator.FromHtml("#b84830"); // terracotta red
    public static readonly Color Black   = ColorTranslator.FromHtml("#0b1008"); // near-black forest
    public static readonly Color TitleBg = ColorTranslator.FromHtml("#0e180a"); // darkest forest
}

// ─── Config file (config.txt next to exe) ───────────────────────────────────
static class Config
{
    static string ConfigFile
    {
        get { return Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "config.txt"); }
    }

    public static void Load(out string src, out string dst)
    {
        src = ""; dst = "";
        if (!File.Exists(ConfigFile)) return;
        foreach (string raw in File.ReadAllLines(ConfigFile))
        {
            string line = raw.Trim();
            if (line.StartsWith("#") || !line.Contains("=")) continue;
            int    idx = line.IndexOf('=');
            string key = line.Substring(0, idx).Trim().ToUpper();
            string val = line.Substring(idx + 1).Trim();
            if (key == "SOURCE")      src = val;
            if (key == "DESTINATION") dst = val;
        }
    }

    public static void Save(string src, string dst)
    {
        File.WriteAllText(ConfigFile,
            "# Robocopy Sync - Default Paths\n" +
            "# Edit these lines, or use the GUI Browse buttons.\n" +
            "# Lines starting with # are ignored.\n\n" +
            "SOURCE="      + src + "\n" +
            "DESTINATION=" + dst + "\n");
    }
}

// ─── Queue item ──────────────────────────────────────────────────────────────
class QueueItem
{
    public int      Id;
    public string   LocalPath;
    public string   ServerPath;
    public string   Direction;   // l2s / s2l / move
    public string   Mode;        // folder / file
    public DateTime ScheduledTime;
    public string   Status;      // Pending / Running / Done / Failed
    public bool     RepeatDaily; // reschedule +1 day after success

    public string DirectionLabel
    {
        get
        {
            if (Direction == "l2s")  return "Local -> Server";
            if (Direction == "s2l")  return "Server -> Local";
            if (Direction == "move") return "Local -> Server + Delete";
            return Direction;
        }
    }
}

// ─── Modern Windows Explorer folder picker via IFileOpenDialog COM ───────────

[ComImport]
[Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IShellItem
{
    void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
    void GetParent(out IShellItem ppsi);
    void GetDisplayName(uint sigdnName,
        [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
    void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
    void Compare(IShellItem psi, uint hint, out int piOrder);
}

[ComImport]
[Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IFileOpenDialog
{
    [PreserveSig] int Show([In] IntPtr parent);
    void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
    void SetFileTypeIndex(uint iFileType);
    void GetFileTypeIndex(out uint piFileType);
    void Advise(IntPtr pfde, out uint pdwCookie);
    void Unadvise(uint dwCookie);
    void SetOptions(uint fos);
    void GetOptions(out uint pfos);
    void SetDefaultFolder(IShellItem psi);
    void SetFolder(IShellItem psi);
    void GetFolder(out IShellItem ppsi);
    void GetCurrentSelection(out IShellItem ppsi);
    void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
    void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
    void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
    void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
    void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
    void GetResult(out IShellItem ppsi);
    void AddPlace(IShellItem psi, int fdap);
    void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
    void Close([MarshalAs(UnmanagedType.Error)] int hr);
    void SetClientGuid(ref Guid guid);
    void ClearClientData();
    void SetFilter(IntPtr pFilter);
    void GetResults(out IntPtr ppenum);
    void GetSelectedItems(out IntPtr ppsai);
}

static class FolderPicker
{
    static readonly Guid CLSID_FileOpenDialog = new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7");
    const uint FOS_PICKFOLDERS     = 0x00000020;
    const uint FOS_FORCEFILESYSTEM = 0x00000040;
    const uint FOS_PATHMUSTEXIST   = 0x00000800;
    const uint SIGDN_FILESYSPATH   = 0x80058000;

    public static string Pick(string title, IntPtr owner)
    {
        IFileOpenDialog dialog = null;
        try
        {
            Type t = Type.GetTypeFromCLSID(CLSID_FileOpenDialog);
            dialog = (IFileOpenDialog)Activator.CreateInstance(t);
            dialog.SetOptions(FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
            dialog.SetTitle(title);
            int hr = dialog.Show(owner);
            if (hr != 0) return null;
            IShellItem item;
            dialog.GetResult(out item);
            string path;
            item.GetDisplayName(SIGDN_FILESYSPATH, out path);
            Marshal.ReleaseComObject(item);
            return path;
        }
        catch { return null; }
        finally
        {
            if (dialog != null) Marshal.ReleaseComObject(dialog);
        }
    }
}

// ─── Main form ───────────────────────────────────────────────────────────────
class MainForm : Form
{
    // ── Sync tab controls ─────────────────────────────────────────────────────
    TextBox          tbLocal, tbServer;
    RadioButton      rbFolder, rbFile;
    Label            lblModeHint, lblStatus, lblCurrentFile;
    RichTextBox      rtbLog;
    Button           btnL2S, btnS2L, btnMove, btnClear, btnOpenLog, btnStop;
    string           currentLogFile;
    ProgressBar      pbFile;
    Label            lblPercent;
    Process          activeProcess;
    Panel            card, btnRow, _syncMain;
    FlowLayoutPanel  bot;
    Label            lblLocalPath, lblServerPath;

    // ── Shared running flag ───────────────────────────────────────────────────
    volatile bool    _isRunning  = false;

    // ── Sync progress (volatile, written by worker, read by UI timer) ─────────
    volatile int    _latestPct      = 0;
    volatile string _latestFile     = "";
    volatile bool   _fileChanged    = false;
    long            _latestFileSize = 0;
    System.Windows.Forms.Timer _progressTimer;

    // ── Queue tab controls ────────────────────────────────────────────────────
    TextBox          tbQLocal, tbQServer;
    RadioButton      rbQDirL2S, rbQDirS2L, rbQDirMove;
    RadioButton      rbQModeFolder, rbQModeFile;
    RadioButton      rbQRunNow, rbQSchedule;
    CheckBox         cbQRepeatDaily;
    DateTimePicker   dtpQSchedule;   // date part
    DateTimePicker   dtpQTime;       // time part
    Button           btnAddToQueue;
    ListView         lvQueue;
    Button           btnQRemove, btnQRunNow, btnQClearAll, btnQStop;
    Label            lblQFile;
    ProgressBar      pbQueue;
    Label            lblQPct;
    RichTextBox      rtbQLog;

    // ── Queue state ───────────────────────────────────────────────────────────
    List<QueueItem>  _queue        = new List<QueueItem>();
    int              _nextQueueId  = 1;
    Process          _queueProcess = null;
    QueueItem        _runningItem  = null;
    System.Windows.Forms.Timer _queueTimer;

    // ── Volatile queue progress (written by queue worker, read by UI) ─────────
    volatile int    _qLatestPct  = 0;
    volatile string _qLatestFile = "";
    volatile bool   _qFileChanged = false;
    long            _qLatestFileSize = 0;
    System.Windows.Forms.Timer _qProgressTimer;

    static string QueueFile
    {
        get { return Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "queue.txt"); }
    }

    public MainForm()
    {
        Text          = "RoboSync";
        MinimumSize   = new Size(960, 720);
        Size          = new Size(1080, 800);
        BackColor     = C.Bg;
        ForeColor     = C.Text;
        Font          = new Font("Segoe UI", 9.5f);
        StartPosition = FormStartPosition.CenterScreen;

        BuildUI();

        string src, dst;
        Config.Load(out src, out dst);
        tbLocal.Text  = src;
        tbServer.Text = dst;

        LoadQueue();

        Shown += delegate {
            if (_syncMain != null) UpdateSyncLayout(_syncMain);
        };
    }

    // ── Build full UI with title + TabControl ─────────────────────────────────
    void BuildUI()
    {
        // Title bar — left-aligned logo + amber bottom border, no centre-paint jitter
        var titlePanel = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = C.TitleBg };

        // Left icon block (solid amber square)
        var iconBlock = new Panel {
            Size = new Size(52, 52), Location = new Point(0, 0),
            BackColor = C.Accent,
        };
        // "RS" monogram inside the block
        var lblMono = new Label {
            Text = "RS", Font = new Font("Segoe UI", 15, FontStyle.Bold),
            ForeColor = C.Black, BackColor = C.Accent,
            TextAlign = ContentAlignment.MiddleCenter,
            Dock = DockStyle.Fill,
        };
        iconBlock.Controls.Add(lblMono);

        // Title text
        var lblTitle = new Label {
            Text = "RoboSync",
            Font = new Font("Segoe UI", 15, FontStyle.Bold),
            ForeColor = C.Text,
            AutoSize = true,
            Location = new Point(64, 8),
        };

        // Subtitle
        var lblSub = new Label {
            Text = "Robocopy File & Folder Sync",
            ForeColor = C.Muted,
            Font = new Font("Segoe UI", 8.5f),
            AutoSize = true,
            Location = new Point(66, 32),
        };

        titlePanel.Controls.Add(iconBlock);
        titlePanel.Controls.Add(lblTitle);
        titlePanel.Controls.Add(lblSub);

        // Amber bottom border line
        titlePanel.Paint += delegate(object s, PaintEventArgs pe) {
            var p = (Panel)s;
            using (var pen = new System.Drawing.Pen(C.Accent, 2))
                pe.Graphics.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);
        };

        Controls.Add(titlePanel);

        // ── Custom tab bar (flat buttons, no TabControl white-strip issue) ───────
        var tabBar = new Panel {
            Dock = DockStyle.Top, Height = 36, BackColor = C.Bg,
        };

        Button btnTabSync  = null;
        Button btnTabQueue = null;
        Panel  syncPanel   = new Panel { Dock = DockStyle.Fill, BackColor = C.Bg, Visible = true  };
        Panel  queuePanel  = new Panel { Dock = DockStyle.Fill, BackColor = C.Bg, Visible = false };

        Action<bool> switchTab = delegate(bool showSync) {
            syncPanel.Visible  =  showSync;
            queuePanel.Visible = !showSync;
            // Repaint both buttons
            if (btnTabSync  != null) btnTabSync.Invalidate();
            if (btnTabQueue != null) btnTabQueue.Invalidate();
        };

        Action<Button, string, bool> makeTab = null;
        makeTab = delegate(Button btn, string label, bool isSync) {
            btn.Text      = label;
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = C.Bg;
            btn.ForeColor = C.Muted;
            btn.Font      = new Font("Segoe UI", 10, FontStyle.Bold);
            btn.Size      = new Size(110, 36);
            btn.Cursor    = Cursors.Hand;
            btn.FlatAppearance.BorderSize       = 0;
            btn.FlatAppearance.MouseOverBackColor = C.Surface;
            bool capturedIsSync = isSync;
            btn.Click += delegate { switchTab(capturedIsSync); };
            btn.Paint += delegate(object s, PaintEventArgs pe) {
                Button b       = (Button)s;
                bool   active  = (capturedIsSync ? syncPanel.Visible : queuePanel.Visible);
                b.ForeColor    = active ? C.Accent : C.Muted;
                if (active)
                {
                    using (var pen = new System.Drawing.Pen(C.Accent, 2))
                        pe.Graphics.DrawLine(pen, 0, b.Height - 2, b.Width, b.Height - 2);
                }
            };
            tabBar.Controls.Add(btn);
        };

        btnTabSync  = new Button();
        btnTabQueue = new Button();
        makeTab(btnTabSync,  "  Sync",  true);
        makeTab(btnTabQueue, "  Queue", false);
        btnTabSync.Location  = new Point(0,   0);
        btnTabQueue.Location = new Point(110, 0);

        // Thin separator line at bottom of tab bar
        tabBar.Paint += delegate(object s, PaintEventArgs pe) {
            var p = (Panel)s;
            using (var pen = new System.Drawing.Pen(C.Overlay, 1))
                pe.Graphics.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);
        };

        // Content area holds both panels (only one visible at a time)
        var contentArea = new Panel { Dock = DockStyle.Fill, BackColor = C.Bg };
        contentArea.Controls.Add(queuePanel);
        contentArea.Controls.Add(syncPanel);

        Controls.Add(contentArea);
        Controls.Add(tabBar);
        tabBar.BringToFront();
        contentArea.BringToFront();

        BuildSyncTab(syncPanel);
        BuildQueueTab(queuePanel);
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SYNC TAB
    // ══════════════════════════════════════════════════════════════════════════
    void BuildSyncTab(Panel page)
    {
        var main = new Panel { Dock = DockStyle.Fill, BackColor = C.Bg,
                               Padding = new Padding(18, 12, 18, 8) };
        page.Controls.Add(main);

        int y = 0;

        // Selection mode
        main.Controls.Add(SectionLabel("Selection Mode", ref y));
        var modeRow = new FlowLayoutPanel { Top = y, Left = 0, Height = 32, AutoSize = true,
                                             BackColor = C.Bg,
                                             FlowDirection = FlowDirection.LeftToRight };
        rbFolder    = Radio("Sync Folder", true,  modeRow);
        rbFile      = Radio("Sync File",   false, modeRow);
        lblModeHint = new Label { Text = "  Browse selects a folder",
                                   ForeColor = C.Muted, AutoSize = true,
                                   TextAlign = ContentAlignment.MiddleLeft };
        modeRow.Controls.Add(lblModeHint);
        rbFolder.CheckedChanged += delegate { UpdateModeHint(); };
        main.Controls.Add(modeRow);
        y += 38;

        // Paths card
        main.Controls.Add(SectionLabel("Paths", ref y));
        card = new Panel { Top = y, Left = 0, BackColor = C.Surface,
                           Height = 90, Padding = new Padding(8, 6, 8, 6) };
        card.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        main.Controls.Add(card);
        tbLocal  = PathRow(card,  0, "Local Path",  BrowseLocal,  out lblLocalPath);
        tbServer = PathRow(card, 42, "Server Path", BrowseServer, out lblServerPath);
        y += 100;

        // Action buttons
        main.Controls.Add(SectionLabel("Actions", ref y));
        btnRow = new Panel { Top = y, Left = 0, Height = 80, BackColor = C.Bg };
        btnRow.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        btnL2S  = ActionBtn("Local  ->  Server\n(copy / mirror)",  C.Accent, "local_to_server");
        btnS2L  = ActionBtn("Server  ->  Local\n(copy / mirror)",  C.Green,  "server_to_local");
        btnMove = ActionBtn("Local  ->  Server\n+ Delete Local",   C.Red,    "move_to_server");
        btnRow.Controls.Add(btnL2S);
        btnRow.Controls.Add(btnS2L);
        btnRow.Controls.Add(btnMove);
        btnRow.Resize += delegate { LayoutActionButtons(); };
        main.Controls.Add(btnRow);
        y += 90;

        // Currently-copying label
        main.Controls.Add(SectionLabel("Log Output", ref y));
        lblCurrentFile = new Label {
            Top = y, Left = 0, Height = 20, AutoEllipsis = true,
            Text = "", ForeColor = C.Muted,
            Font = new Font("Consolas", 8.5f),
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        main.Controls.Add(lblCurrentFile);
        y += 22;

        // Per-file progress bar
        pbFile = new ProgressBar {
            Top = y, Left = 0, Height = 14, Width = 900,
            Minimum = 0, Maximum = 100, Value = 0,
            Style = ProgressBarStyle.Continuous,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        lblPercent = new Label {
            Top = y - 1, Left = 0, Width = 46, Height = 16,
            Text = "", ForeColor = C.Muted,
            Font = new Font("Consolas", 8f),
            TextAlign = ContentAlignment.MiddleRight,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
        };
        main.Controls.Add(pbFile);
        main.Controls.Add(lblPercent);
        y += 20;

        // Log
        rtbLog = new RichTextBox {
            Top = y, Left = 0, Height = 230,
            BackColor = C.Black, ForeColor = C.Text,
            Font = new Font("Consolas", 9),
            ReadOnly = true, BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.Both, WordWrap = false,
        };
        rtbLog.Anchor = AnchorStyles.Left | AnchorStyles.Right |
                        AnchorStyles.Top  | AnchorStyles.Bottom;
        main.Controls.Add(rtbLog);
        y += 238;

        // Bottom bar
        bot = new FlowLayoutPanel { Top = y, Left = 0, Height = 32, AutoSize = true,
                                    BackColor = C.Bg,
                                    FlowDirection = FlowDirection.LeftToRight };
        bot.Anchor = AnchorStyles.Left | AnchorStyles.Bottom;
        btnClear   = SmallBtn("Clear Log",     C.Muted,  "clear");
        btnOpenLog = SmallBtn("Open Log File", C.Accent, "openlog");
        btnStop    = SmallBtn("Stop",          C.Red,    "stop");
        btnOpenLog.Enabled = false;
        btnStop.Enabled    = false;
        bot.Controls.Add(btnClear);
        bot.Controls.Add(btnOpenLog);
        bot.Controls.Add(btnStop);
        main.Controls.Add(bot);

        lblStatus = new Label { Text = "Ready", ForeColor = C.Muted, AutoSize = true };
        lblStatus.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
        main.Controls.Add(lblStatus);

        _syncMain = main;
        main.Resize += delegate { UpdateSyncLayout(main); };

        // 80 ms UI refresh — reads volatile fields written by the worker thread
        _progressTimer = new System.Windows.Forms.Timer { Interval = 80 };
        _progressTimer.Tick += delegate
        {
            int    p = _latestPct;
            string f = _latestFile;

            if (p < 100) { pbFile.Value = p + 1; pbFile.Value = p; }
            else         { pbFile.Value = 100; }
            lblPercent.Text = p + "%";

            if (_fileChanged)
            {
                _fileChanged = false;
            }
            if (f.Length > 0)
            {
                long size        = System.Threading.Interlocked.Read(ref _latestFileSize);
                long transferred = size > 0 ? (long)(size * (p / 100.0)) : 0;
                string sizeInfo  = size > 0
                    ? "   " + FormatBytes(transferred) + " / " + FormatBytes(size)
                    : "";
                lblCurrentFile.Text = "  Copying: " + f + "  [" + p + "%]" + sizeInfo;
            }
        };

        // Queue 80 ms progress timer
        _qProgressTimer = new System.Windows.Forms.Timer { Interval = 80 };
        _qProgressTimer.Tick += delegate
        {
            int    p = _qLatestPct;
            string f = _qLatestFile;

            if (p < 100) { pbQueue.Value = p + 1; pbQueue.Value = p; }
            else         { pbQueue.Value = 100; }
            lblQPct.Text = p + "%";

            if (_qFileChanged)
            {
                _qFileChanged = false;
            }
            if (f.Length > 0)
            {
                long size        = System.Threading.Interlocked.Read(ref _qLatestFileSize);
                long transferred = size > 0 ? (long)(size * (p / 100.0)) : 0;
                string sizeInfo  = size > 0
                    ? "   " + FormatBytes(transferred) + " / " + FormatBytes(size)
                    : "";
                lblQFile.Text = "  Copying: " + f + "  [" + p + "%]" + sizeInfo;
            }
        };

        // Queue scheduler timer — fires every 15 seconds
        _queueTimer = new System.Windows.Forms.Timer { Interval = 15000 };
        _queueTimer.Tick += delegate { CheckQueueScheduler(); };
        _queueTimer.Start();
    }

    void UpdateSyncLayout(Panel main)
    {
        int w = main.ClientSize.Width - main.Padding.Left - main.Padding.Right;
        if (w < 10) return;
        if (card          != null) card.Width           = w;
        if (btnRow        != null) btnRow.Width         = w;
        if (lblCurrentFile!= null) lblCurrentFile.Width = w;
        if (pbFile        != null) pbFile.Width         = w - 54;
        if (rtbLog        != null) rtbLog.Width         = w;
        if (card          != null) card.PerformLayout();
        LayoutActionButtons();
        if (lblStatus != null && bot != null)
            lblStatus.Location = new Point(w - lblStatus.Width, bot.Top + 6);
    }

    void LayoutActionButtons()
    {
        if (btnRow == null || btnL2S == null) return;
        int w = btnRow.ClientSize.Width;
        int h = btnRow.ClientSize.Height;
        if (w < 1 || h < 1) return;

        bool fileMode = rbFile != null && rbFile.Checked;
        if (fileMode)
        {
            int half = w / 2;
            btnL2S.SetBounds(0,        0, half - 4, h);
            btnMove.SetBounds(half + 4, 0, w - half - 4, h);
            btnS2L.Visible = false;
            btnL2S.Text    = "File  ->  Folder\n(copy)";
            btnMove.Text   = "File  ->  Folder\n+ Delete File";
        }
        else
        {
            int third = w / 3;
            btnL2S.SetBounds(0,          0, third - 4,          h);
            btnS2L.SetBounds(third + 2,  0, third - 4,          h);
            btnMove.SetBounds(third*2+4, 0, w - (third*2+4),    h);
            btnS2L.Visible = true;
            btnL2S.Text    = "Local  ->  Server\n(copy / mirror)";
            btnMove.Text   = "Local  ->  Server\n+ Delete Local";
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // QUEUE TAB
    // ══════════════════════════════════════════════════════════════════════════
    void BuildQueueTab(Panel page)
    {
        var scroll = new Panel { Dock = DockStyle.Fill, BackColor = C.Bg,
                                  AutoScroll = true, Padding = new Padding(18, 12, 18, 8) };
        page.Controls.Add(scroll);

        int y = 0;

        // ── Add to Queue section ──────────────────────────────────────────────
        scroll.Controls.Add(QSectionLabel("Add to Queue", ref y));

        var addCard = new Panel { Top = y, Left = 0, BackColor = C.Surface,
                                   Height = 260, Padding = new Padding(10, 8, 10, 8) };
        addCard.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        scroll.Controls.Add(addCard);

        // Local path row
        var lblQL = new Label { Text = "Local Path", ForeColor = C.Text, AutoSize = true,
                                 Top = 8, Left = 4,
                                 Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        addCard.Controls.Add(lblQL);
        tbQLocal = new TextBox { Top = 6, Left = 110, Height = 26, Width = 400,
                                  BackColor = C.Overlay, ForeColor = C.Text,
                                  BorderStyle = BorderStyle.None,
                                  Font = new Font("Segoe UI", 10) };
        tbQLocal.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        addCard.Controls.Add(tbQLocal);
        var btnQBrowseLocal = new Button { Text = "Browse", Top = 6, Width = 70, Height = 26,
                                            BackColor = C.Overlay, ForeColor = C.Accent,
                                            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnQBrowseLocal.FlatAppearance.BorderSize = 0;
        btnQBrowseLocal.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnQBrowseLocal.Click += delegate { QBrowseLocal(); };
        addCard.Controls.Add(btnQBrowseLocal);

        // Server path row
        var lblQS = new Label { Text = "Server Path", ForeColor = C.Text, AutoSize = true,
                                 Top = 42, Left = 4,
                                 Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        addCard.Controls.Add(lblQS);
        tbQServer = new TextBox { Top = 40, Left = 110, Height = 26, Width = 400,
                                   BackColor = C.Overlay, ForeColor = C.Text,
                                   BorderStyle = BorderStyle.None,
                                   Font = new Font("Segoe UI", 10) };
        tbQServer.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        addCard.Controls.Add(tbQServer);
        var btnQBrowseServer = new Button { Text = "Browse", Top = 40, Width = 70, Height = 26,
                                             BackColor = C.Overlay, ForeColor = C.Accent,
                                             FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btnQBrowseServer.FlatAppearance.BorderSize = 0;
        btnQBrowseServer.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btnQBrowseServer.Click += delegate { QBrowseServer(); };
        addCard.Controls.Add(btnQBrowseServer);

        addCard.Resize += delegate {
            tbQLocal.Width        = addCard.ClientSize.Width - 110 - 80 - 12;
            btnQBrowseLocal.Left  = addCard.ClientSize.Width - 80 - 4;
            tbQServer.Width       = addCard.ClientSize.Width - 110 - 80 - 12;
            btnQBrowseServer.Left = addCard.ClientSize.Width - 80 - 4;
        };

        // Direction radios
        var lblQDir = new Label { Text = "DIRECTION", ForeColor = C.Muted, AutoSize = true,
                                   Top = 78, Left = 4,
                                   Font = new Font("Segoe UI", 7.5f, FontStyle.Bold) };
        addCard.Controls.Add(lblQDir);
        var dirRow = new FlowLayoutPanel { Top = 94, Left = 4, Height = 28, AutoSize = true,
                                            BackColor = C.Surface,
                                            FlowDirection = FlowDirection.LeftToRight };
        rbQDirL2S  = RadioQ("Local → Server",          true,  dirRow);
        rbQDirS2L  = RadioQ("Server → Local",          false, dirRow);
        rbQDirMove = RadioQ("Local → Server + Delete", false, dirRow);
        addCard.Controls.Add(dirRow);

        // Mode radios
        var lblQMode = new Label { Text = "MODE", ForeColor = C.Muted, AutoSize = true,
                                    Top = 130, Left = 4,
                                    Font = new Font("Segoe UI", 7.5f, FontStyle.Bold) };
        addCard.Controls.Add(lblQMode);
        var modeQRow = new FlowLayoutPanel { Top = 146, Left = 4, Height = 28, AutoSize = true,
                                              BackColor = C.Surface,
                                              FlowDirection = FlowDirection.LeftToRight };
        rbQModeFolder = RadioQ("Folder", true,  modeQRow);
        rbQModeFile   = RadioQ("File",   false, modeQRow);
        addCard.Controls.Add(modeQRow);

        // Schedule section
        var lblQSched = new Label { Text = "SCHEDULE", ForeColor = C.Muted, AutoSize = true,
                                     Top = 182, Left = 4,
                                     Font = new Font("Segoe UI", 7.5f, FontStyle.Bold) };
        addCard.Controls.Add(lblQSched);
        var schedRow = new FlowLayoutPanel { Top = 198, Left = 4, Height = 30, AutoSize = true,
                                              BackColor = C.Surface,
                                              FlowDirection = FlowDirection.LeftToRight };
        // Radio: Run Now (default)
        rbQRunNow = new RadioButton {
            Text    = "Run Now", Checked = true,
            ForeColor = C.Text, BackColor = C.Surface,
            AutoSize = true, Padding = new Padding(0, 0, 14, 0),
        };
        // Radio: Schedule
        rbQSchedule = new RadioButton {
            Text    = "Schedule", Checked = false,
            ForeColor = C.Accent, BackColor = C.Surface,
            AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            Padding = new Padding(0, 0, 10, 0),
        };
        // Date picker
        dtpQSchedule = new DateTimePicker {
            Format       = DateTimePickerFormat.Custom,
            CustomFormat = "yyyy-MM-dd",
            Value        = DateTime.Now.AddMinutes(30),
            Width        = 130, Height = 26,
            BackColor    = C.Overlay, ForeColor = C.Text,
            CalendarForeColor = C.Text, CalendarMonthBackground = C.Surface,
            Visible = false,
        };
        // Time picker
        dtpQTime = new DateTimePicker {
            Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm",
            ShowUpDown = true, Value = DateTime.Now.AddMinutes(30),
            Width = 72, Height = 26,
            BackColor = C.Overlay, ForeColor = C.Text,
            Visible = false,
        };
        cbQRepeatDaily = new CheckBox {
            Text = "Repeat every day", Checked = false, Visible = false,
            ForeColor = C.Accent, BackColor = C.Surface,
            AutoSize = true, Padding = new Padding(10, 0, 0, 0),
        };
        rbQSchedule.CheckedChanged += delegate {
            bool scheduled = rbQSchedule.Checked;
            dtpQSchedule.Visible   = scheduled;
            dtpQTime.Visible       = scheduled;
            cbQRepeatDaily.Visible = scheduled;
            if (!scheduled) cbQRepeatDaily.Checked = false;
        };
        schedRow.Controls.Add(rbQRunNow);
        schedRow.Controls.Add(rbQSchedule);
        schedRow.Controls.Add(dtpQSchedule);
        schedRow.Controls.Add(dtpQTime);
        schedRow.Controls.Add(cbQRepeatDaily);
        addCard.Controls.Add(schedRow);

        // Add to Queue button
        btnAddToQueue = new Button {
            Text      = "Add to Queue",
            Top       = 228, Left = 4,
            Width     = 160, Height = 28,
            BackColor = C.Accent,
            ForeColor = C.Black,
            FlatStyle = FlatStyle.Flat,
            Font      = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor    = Cursors.Hand,
        };
        btnAddToQueue.FlatAppearance.BorderSize = 0;
        btnAddToQueue.Click += delegate { AddToQueue(); };
        addCard.Controls.Add(btnAddToQueue);

        y += 270;

        // ── Scheduled Queue section ───────────────────────────────────────────
        scroll.Controls.Add(QSectionLabel("Scheduled Queue", ref y));

        // Queue ListView
        lvQueue = new ListView {
            Top        = y, Left = 0, Height = 160,
            View       = View.Details,
            FullRowSelect = true,
            GridLines  = true,
            BackColor  = C.Black,
            ForeColor  = C.Text,
            BorderStyle = BorderStyle.None,
            Font       = new Font("Consolas", 9),
        };
        lvQueue.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        lvQueue.Columns.Add("#",           36);
        lvQueue.Columns.Add("Source",     180);
        lvQueue.Columns.Add("Destination",180);
        lvQueue.Columns.Add("Direction",  130);
        lvQueue.Columns.Add("Mode",        60);
        lvQueue.Columns.Add("Scheduled",  140);
        lvQueue.Columns.Add("Status",      80);
        scroll.Controls.Add(lvQueue);
        y += 168;

        // Queue action buttons
        var qBtnRow = new FlowLayoutPanel { Top = y, Left = 0, Height = 32, AutoSize = true,
                                             BackColor = C.Bg,
                                             FlowDirection = FlowDirection.LeftToRight };
        qBtnRow.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        btnQRemove  = QSmallBtn("Remove Selected", C.Red);
        btnQRunNow  = QSmallBtn("Run Selected Now", C.Green);
        btnQClearAll= QSmallBtn("Clear All",        C.Muted);
        btnQStop    = QSmallBtn("Stop",             C.Red);
        btnQStop.Enabled = false;
        btnQRemove.Click   += delegate { QueueRemoveSelected(); };
        btnQRunNow.Click   += delegate { QueueRunSelectedNow(); };
        btnQClearAll.Click += delegate { QueueClearAll(); };
        btnQStop.Click     += delegate { QueueStop(); };
        qBtnRow.Controls.Add(btnQRemove);
        qBtnRow.Controls.Add(btnQRunNow);
        qBtnRow.Controls.Add(btnQClearAll);
        qBtnRow.Controls.Add(btnQStop);
        scroll.Controls.Add(qBtnRow);
        y += 38;

        // ── Queue log section ─────────────────────────────────────────────────
        scroll.Controls.Add(QSectionLabel("Queue Transfer Log", ref y));

        lblQFile = new Label {
            Top = y, Left = 0, Height = 20, AutoEllipsis = true,
            Text = "", ForeColor = C.Muted,
            Font = new Font("Consolas", 8.5f),
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        scroll.Controls.Add(lblQFile);
        y += 22;

        pbQueue = new ProgressBar {
            Top = y, Left = 0, Height = 14, Width = 900,
            Minimum = 0, Maximum = 100, Value = 0,
            Style = ProgressBarStyle.Continuous,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        lblQPct = new Label {
            Top = y - 1, Left = 0, Width = 46, Height = 16,
            Text = "", ForeColor = C.Muted,
            Font = new Font("Consolas", 8f),
            TextAlign = ContentAlignment.MiddleRight,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
        };
        scroll.Controls.Add(pbQueue);
        scroll.Controls.Add(lblQPct);
        y += 20;

        rtbQLog = new RichTextBox {
            Top = y, Left = 0, Height = 200,
            BackColor = C.Black, ForeColor = C.Text,
            Font = new Font("Consolas", 9),
            ReadOnly = true, BorderStyle = BorderStyle.None,
            ScrollBars = RichTextBoxScrollBars.Both, WordWrap = false,
        };
        rtbQLog.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
        scroll.Controls.Add(rtbQLog);

        scroll.Resize += delegate { UpdateQueueLayout(scroll); };
    }

    void UpdateQueueLayout(Panel scroll)
    {
        int w = scroll.ClientSize.Width - scroll.Padding.Horizontal;
        if (w < 1) return;
        if (lvQueue  != null) lvQueue.Width  = w;
        if (lblQFile != null) lblQFile.Width = w;
        if (pbQueue  != null) pbQueue.Width  = w - 54;
        if (rtbQLog  != null) rtbQLog.Width  = w;
    }

    Label QSectionLabel(string text, ref int y)
    {
        var lbl = new Label { Text = text.ToUpper(), ForeColor = C.Muted, AutoSize = true,
                               Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                               Top = y, Left = 0 };
        y += 22;
        return lbl;
    }

    RadioButton RadioQ(string text, bool isChecked, Control parent)
    {
        var rb = new RadioButton { Text = text, Checked = isChecked, AutoSize = true,
                                    ForeColor = C.Text, BackColor = C.Surface,
                                    Padding = new Padding(4, 0, 12, 0) };
        parent.Controls.Add(rb);
        return rb;
    }

    Button QSmallBtn(string text, Color fg)
    {
        var b = new Button {
            Text = text, ForeColor = fg, BackColor = C.Overlay,
            AutoSize = true, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 6, 0), Padding = new Padding(8, 2, 8, 2),
        };
        b.FlatAppearance.BorderSize = 0;
        return b;
    }

    // ── Queue browse helpers ──────────────────────────────────────────────────
    void QBrowseLocal()
    {
        if (rbQModeFolder.Checked)
        {
            string p = FolderPicker.Pick("Select Local Folder", Handle);
            if (p != null) tbQLocal.Text = p;
        }
        else
        {
            using (var d = new OpenFileDialog { Title = "Select Source File" })
                if (d.ShowDialog() == DialogResult.OK) tbQLocal.Text = d.FileName;
        }
    }

    void QBrowseServer()
    {
        string p = FolderPicker.Pick("Select Server / Destination Folder", Handle);
        if (p != null) tbQServer.Text = p;
    }

    // ── Add item to queue ─────────────────────────────────────────────────────
    void AddToQueue()
    {
        string local  = tbQLocal.Text.Trim();
        string server = tbQServer.Text.Trim();
        if (string.IsNullOrEmpty(local) || string.IsNullOrEmpty(server))
        {
            MessageBox.Show("Please enter both Local and Server paths.", "Missing Paths",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string dir = "l2s";
        if (rbQDirS2L.Checked)  dir = "s2l";
        if (rbQDirMove.Checked) dir = "move";

        string mode = rbQModeFile.Checked ? "file" : "folder";

        DateTime sched = rbQSchedule.Checked
            ? dtpQSchedule.Value.Date
                .AddHours(dtpQTime.Value.Hour)
                .AddMinutes(dtpQTime.Value.Minute)
            : DateTime.Now;

        var item = new QueueItem {
            Id            = _nextQueueId++,
            LocalPath     = local,
            ServerPath    = server,
            Direction     = dir,
            Mode          = mode,
            ScheduledTime = sched,
            Status        = "Pending",
            RepeatDaily   = cbQRepeatDaily.Checked,
        };

        _queue.Add(item);
        SaveQueue();
        RefreshQueueListView();

        // Clear inputs
        tbQLocal.Text  = "";
        tbQServer.Text = "";
        rbQRunNow.Checked = true;
    }

    // ── Queue ListView refresh ────────────────────────────────────────────────
    void RefreshQueueListView()
    {
        lvQueue.Items.Clear();
        int rowNum = 1;
        foreach (QueueItem item in _queue)
        {
            string schedStr = (item.ScheduledTime <= DateTime.Now && item.Status == "Pending"
                ? "ASAP"
                : item.ScheduledTime.ToString("yyyy-MM-dd HH:mm"))
                + (item.RepeatDaily ? "  [Daily]" : "");

            var lvi = new ListViewItem((rowNum++).ToString());
            lvi.SubItems.Add(item.LocalPath);
            lvi.SubItems.Add(item.ServerPath);
            lvi.SubItems.Add(item.DirectionLabel);
            lvi.SubItems.Add(item.Mode);
            lvi.SubItems.Add(schedStr);
            lvi.SubItems.Add(item.Status);
            lvi.Tag = item;

            Color statusColor;
            if      (item.Status == "Running") statusColor = C.Yellow;
            else if (item.Status == "Done")    statusColor = C.Green;
            else if (item.Status == "Failed")  statusColor = C.Red;
            else                               statusColor = C.Text;

            lvi.ForeColor = statusColor;
            lvQueue.Items.Add(lvi);
        }
    }

    // ── Queue persistence ─────────────────────────────────────────────────────
    void SaveQueue()
    {
        try
        {
            var lines = new List<string>();
            foreach (QueueItem item in _queue)
            {
                lines.Add(item.LocalPath + "|" +
                           item.ServerPath + "|" +
                           item.Direction + "|" +
                           item.Mode + "|" +
                           item.ScheduledTime.ToString("yyyy-MM-dd HH:mm:ss") + "|" +
                           item.Status + "|" +
                           (item.RepeatDaily ? "1" : "0"));
            }
            File.WriteAllLines(QueueFile, lines.ToArray());
        }
        catch { }
    }

    void LoadQueue()
    {
        if (!File.Exists(QueueFile)) return;
        try
        {
            foreach (string raw in File.ReadAllLines(QueueFile))
            {
                string line = raw.Trim();
                if (line.Length == 0) continue;
                string[] parts = line.Split('|');
                if (parts.Length < 6) continue;
                string status = parts[5].Trim();
                if (status == "Done" || status == "Failed") continue;
                if (status == "Running") status = "Pending";

                DateTime dt;
                if (!DateTime.TryParse(parts[4].Trim(), out dt)) dt = DateTime.Now;

                var item = new QueueItem {
                    Id            = _nextQueueId++,
                    LocalPath     = parts[0].Trim(),
                    ServerPath    = parts[1].Trim(),
                    Direction     = parts[2].Trim(),
                    Mode          = parts[3].Trim(),
                    ScheduledTime = dt,
                    Status        = status,
                    RepeatDaily   = parts.Length > 6 && parts[6].Trim() == "1",
                };
                _queue.Add(item);
            }
            RefreshQueueListView();
        }
        catch { }
    }

    // ── Queue scheduler ───────────────────────────────────────────────────────
    void CheckQueueScheduler()
    {
        if (_isRunning) return;

        foreach (QueueItem item in _queue)
        {
            if (item.Status == "Pending" && item.ScheduledTime <= DateTime.Now)
            {
                RunQueueItem(item);
                return;
            }
        }
    }

    // ── Queue actions ─────────────────────────────────────────────────────────
    void QueueRemoveSelected()
    {
        if (lvQueue.SelectedItems.Count == 0) return;
        foreach (ListViewItem lvi in lvQueue.SelectedItems)
        {
            QueueItem item = (QueueItem)lvi.Tag;
            if (item.Status == "Running") continue;
            _queue.Remove(item);
        }
        SaveQueue();
        RefreshQueueListView();
    }

    void QueueRunSelectedNow()
    {
        if (_isRunning)
        {
            MessageBox.Show("A sync is already running. Please wait.", "Busy",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        if (lvQueue.SelectedItems.Count == 0) return;
        QueueItem item = (QueueItem)lvQueue.SelectedItems[0].Tag;
        if (item.Status == "Running") return;
        item.ScheduledTime = DateTime.Now;
        item.Status = "Pending";
        RunQueueItem(item);
    }

    void QueueClearAll()
    {
        var toRemove = new List<QueueItem>();
        foreach (QueueItem item in _queue)
        {
            if (item.Status != "Running") toRemove.Add(item);
        }
        foreach (QueueItem item in toRemove) _queue.Remove(item);
        SaveQueue();
        RefreshQueueListView();
    }

    void QueueStop()
    {
        Process proc = _queueProcess;
        if (proc == null || proc.HasExited) return;
        try { proc.Kill(); } catch { }

        _qProgressTimer.Stop();
        btnQStop.Enabled = false;
        pbQueue.Value    = 0;
        lblQPct.Text     = "";

        string divider = new string('-', 60);
        AppendQLog("\n" + divider + "\n", C.Muted);
        AppendQLog("  STOPPED by user.\n", C.Red);
        AppendQLog(divider + "\n", C.Muted);

        if (_runningItem != null)
        {
            _runningItem.Status = "Failed";
            _runningItem = null;
        }

        _queueProcess = null;
        _isRunning    = false;
        RefreshQueueListView();
        SaveQueue();
        SetSyncButtonsEnabled(true);
        SetQueueButtonsEnabled(true);
    }

    // ── Run a queue item ──────────────────────────────────────────────────────
    void RunQueueItem(QueueItem item)
    {
        if (_isRunning) return;

        _isRunning    = true;
        _runningItem  = item;
        item.Status   = "Running";
        RefreshQueueListView();
        SaveQueue();

        SetSyncButtonsEnabled(false);
        SetQueueButtonsEnabled(false);
        btnQStop.Enabled = true;

        string ts      = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string logDir  = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "log");
        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
        string logFile = Path.Combine(logDir, "queue_log_" + ts + ".txt");

        string args;
        bool isFolder = (item.Mode == "folder");

        if (isFolder)
        {
            string src = (item.Direction != "s2l") ? item.LocalPath  : item.ServerPath;
            string dst = (item.Direction != "s2l") ? item.ServerPath : item.LocalPath;

            if (item.Direction == "move")
                args = Quote(src) + " " + Quote(dst) + " /E /MOVE /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
            else
                args = Quote(src) + " " + Quote(dst) + " /MIR /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
        }
        else
        {
            string srcDir, srcFile, dstDir;
            if (item.Direction == "l2s")
            {
                srcDir  = Path.GetDirectoryName(item.LocalPath);
                srcFile = Path.GetFileName(item.LocalPath);
                dstDir  = item.ServerPath;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
            }
            else if (item.Direction == "s2l")
            {
                srcDir  = Path.GetDirectoryName(item.ServerPath);
                srcFile = Path.GetFileName(item.ServerPath);
                dstDir  = item.LocalPath;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
            }
            else
            {
                srcDir  = Path.GetDirectoryName(item.LocalPath);
                srcFile = Path.GetFileName(item.LocalPath);
                dstDir  = item.ServerPath;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /MOV /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
            }
        }

        _qLatestPct       = 0;
        _qLatestFile      = "";
        _qFileChanged     = false;
        System.Threading.Interlocked.Exchange(ref _qLatestFileSize, 0);
        pbQueue.Value     = 0;
        lblQPct.Text      = "0%";
        lblQFile.Text     = "";
        _qProgressTimer.Start();

        string divider = new string('-', 60);
        AppendQLog("\n" + divider + "\n", C.Muted);
        AppendQLog("  Queue Item #" + item.Id + "  " + item.DirectionLabel + "\n", C.Accent);
        AppendQLog("  " + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss") + "\n", C.Muted);
        AppendQLog(divider + "\n\n", C.Muted);

        string capturedArgs = args;
        string capturedLog  = logFile;
        QueueItem capturedItem = item;
        ThreadPool.QueueUserWorkItem(delegate { QueueWorker(capturedArgs, capturedLog, capturedItem); });
    }

    // ── Queue worker thread ───────────────────────────────────────────────────
    void QueueWorker(string args, string logFile, QueueItem item)
    {
        try
        {
            var psi = new ProcessStartInfo("robocopy", args) {
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
            };

            using (var proc = Process.Start(psi))
            {
                BeginInvoke((Action)delegate { _queueProcess = proc; });

                proc.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e)
                {
                    if (e.Data == null) return;
                    string raw = e.Data;
                    string t   = raw.TrimEnd('\r').Trim();
                    if (t.Length == 0) return;

                    if (t.EndsWith("%"))
                    {
                        string digits = t.TrimEnd('%').Trim();
                        int pct;
                        if (int.TryParse(digits, out pct) && pct >= 0 && pct <= 100)
                            _qLatestPct = pct;
                        return;
                    }

                    string filename, sizeRaw;
                    ExtractNameAndSize(raw, out filename, out sizeRaw);
                    if (filename != null)
                    {
                        _qLatestPct  = 0;
                        _qLatestFile = filename;
                        System.Threading.Interlocked.Exchange(ref _qLatestFileSize, ParseRoboSize(sizeRaw));
                        _qFileChanged = true;
                    }

                    string captured = raw;
                    BeginInvoke((Action)delegate { ProcessQueueLine(captured); });
                };

                proc.BeginOutputReadLine();
                proc.WaitForExit();
                int rc = proc.ExitCode;
                BeginInvoke((Action)delegate { FinishQueue(rc, logFile, item); });
            }
        }
        catch (Exception ex)
        {
            string msg = ex.Message;
            BeginInvoke((Action)delegate {
                AppendQLog("\nError: " + msg + "\n", C.Red);
                FinishQueue(-1, logFile, item);
            });
        }
    }

    // ── Parse robocopy lines for queue log ────────────────────────────────────
    void ProcessQueueLine(string line)
    {
        string t = line.Trim();
        if (t.Length == 0) { AppendQLog("\n", C.Text); return; }
        if (t.EndsWith("%") && !t.Contains(" ")) return;

        Color color = C.Text;

        if (t.IndexOf("New File",  StringComparison.OrdinalIgnoreCase) >= 0 ||
            t.IndexOf("New Dir",   StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Green;
        else if (t.IndexOf("Newer",   StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("Older",   StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("Changed", StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Yellow;
        else if (t.IndexOf("*EXTRA",   StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("*Skipping",StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Muted;
        else if (t.IndexOf("FAILED", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("ERROR",  StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Red;
        else if (t.StartsWith("Dirs :") || t.StartsWith("Files :") ||
                 t.StartsWith("Bytes :") || t.StartsWith("Times :"))
            color = C.Accent;
        else if (t.StartsWith("---") || t.StartsWith("ROBOCOPY") ||
                 t.StartsWith("Started") || t.StartsWith("Ended") ||
                 t.StartsWith("Source") || t.StartsWith("Dest") ||
                 t.StartsWith("Options"))
            color = C.Muted;

        AppendQLog(line + "\n", color);
    }

    // ── Finish queue item ─────────────────────────────────────────────────────
    void FinishQueue(int rc, string logFile, QueueItem item)
    {
        _qProgressTimer.Stop();
        _queueProcess    = null;
        btnQStop.Enabled = false;
        lblQFile.Text    = "";
        pbQueue.Value    = rc <= 7 ? 100 : 0;
        lblQPct.Text     = rc <= 7 ? "100%" : "";
        AppendQLog("\n");

        if (rc <= 7)
        {
            if (item.RepeatDaily)
            {
                // Reschedule for the same time tomorrow
                item.ScheduledTime = item.ScheduledTime.AddDays(1);
                item.Status        = "Pending";
                AppendQLog(rc <= 1 ? "  Completed successfully.\n" :
                           rc <= 3 ? "  Completed - some files copied.\n" :
                                     "  Completed with warnings.\n", C.Green);
                AppendQLog("  Rescheduled daily -> next run: " +
                           item.ScheduledTime.ToString("yyyy-MM-dd HH:mm") + "\n", C.Accent);
            }
            else
            {
                item.Status = "Done";
                AppendQLog(rc <= 1 ? "  Completed successfully.\n" :
                           rc <= 3 ? "  Completed - some files copied.\n" :
                                     "  Completed with warnings.\n", C.Green);
            }
        }
        else
        {
            item.Status = "Failed";
            AppendQLog("  Errors (exit code " + rc + "). Check the log.\n", C.Red);
        }
        AppendQLog("\n  Log: " + logFile + "\n", C.Muted);

        _runningItem  = null;
        _isRunning    = false;
        RefreshQueueListView();
        SaveQueue();
        SetSyncButtonsEnabled(true);
        SetQueueButtonsEnabled(true);
    }

    // ── Queue log helper ──────────────────────────────────────────────────────
    void AppendQLog(string text, Color? color = null)
    {
        rtbQLog.SelectionStart  = rtbQLog.TextLength;
        rtbQLog.SelectionLength = 0;
        rtbQLog.SelectionColor  = color.HasValue ? color.Value : C.Text;
        rtbQLog.AppendText(text);
        rtbQLog.ScrollToCaret();
    }

    void SetQueueButtonsEnabled(bool enabled)
    {
        btnAddToQueue.Enabled = enabled;
        btnQRemove.Enabled    = enabled;
        btnQRunNow.Enabled    = enabled;
        btnQClearAll.Enabled  = enabled;
    }

    // ══════════════════════════════════════════════════════════════════════════
    // SHARED HELPERS
    // ══════════════════════════════════════════════════════════════════════════

    Label SectionLabel(string text, ref int y)
    {
        var lbl = new Label { Text = text.ToUpper(), ForeColor = C.Muted, AutoSize = true,
                               Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                               Top = y, Left = 0 };
        y += 22;
        return lbl;
    }

    RadioButton Radio(string text, bool isChecked, Control parent)
    {
        var rb = new RadioButton { Text = text, Checked = isChecked, AutoSize = true,
                                    ForeColor = C.Text, BackColor = C.Bg,
                                    Padding = new Padding(4, 0, 12, 0) };
        parent.Controls.Add(rb);
        return rb;
    }

    TextBox PathRow(Panel owner, int top, string label, Action browseAction, out Label outLabel)
    {
        var lbl = new Label { Text = label, ForeColor = C.Text, AutoSize = true,
                               Top = top + 8, Left = 4,
                               Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        outLabel = lbl;
        var tb  = new TextBox { Top = top + 4, Left = 110, Height = 26, Width = 520,
                                 BackColor = C.Overlay, ForeColor = C.Text,
                                 BorderStyle = BorderStyle.None,
                                 Font = new Font("Segoe UI", 10) };
        tb.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;

        Action capturedBrowse = browseAction;
        var btn = new Button { Text = "Browse", Top = top + 4, Width = 70, Height = 26,
                                BackColor = C.Overlay, ForeColor = C.Accent,
                                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
        btn.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        btn.FlatAppearance.BorderSize = 0;
        btn.Click += delegate { capturedBrowse(); };

        TextBox capturedTb  = tb;
        Button  capturedBtn = btn;
        owner.Resize += delegate {
            capturedTb.Width = owner.ClientSize.Width - 110 - 80 - 12;
            capturedBtn.Left = owner.ClientSize.Width - 80 - 4;
        };
        owner.Controls.Add(lbl);
        owner.Controls.Add(tb);
        owner.Controls.Add(btn);
        return tb;
    }

    Button ActionBtn(string text, Color fg, string mode)
    {
        string capturedMode = mode;
        var b = new Button {
            Text = text, ForeColor = fg, BackColor = C.Overlay,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
        };
        b.FlatAppearance.BorderSize = 0;
        b.Click += delegate { RunSync(capturedMode); };
        return b;
    }

    Button SmallBtn(string text, Color fg, string action)
    {
        string capturedAction = action;
        var b = new Button {
            Text = text, ForeColor = fg, BackColor = C.Overlay,
            AutoSize = true, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 6, 0), Padding = new Padding(8, 2, 8, 2),
        };
        b.FlatAppearance.BorderSize = 0;
        b.Click += delegate {
            if (capturedAction == "clear")   rtbLog.Clear();
            if (capturedAction == "openlog") OpenLog();
            if (capturedAction == "stop")    StopSync();
        };
        return b;
    }

    void UpdateModeHint()
    {
        bool fileMode = rbFile.Checked;
        lblModeHint.Text    = fileMode
            ? "  Browse selects a file / folder"
            : "  Browse selects a folder";
        lblLocalPath.Text   = fileMode ? "File"   : "Local Path";
        lblServerPath.Text  = fileMode ? "Folder" : "Server Path";
        LayoutActionButtons();
    }

    // ── Browse dialogs ────────────────────────────────────────────────────────
    void BrowseLocal()
    {
        if (rbFolder.Checked)
        {
            string p = FolderPicker.Pick("Select Local Folder", Handle);
            if (p != null) tbLocal.Text = p;
        }
        else
        {
            using (var d = new OpenFileDialog { Title = "Select Source File" })
                if (d.ShowDialog() == DialogResult.OK) tbLocal.Text = d.FileName;
        }
    }

    void BrowseServer()
    {
        string p = FolderPicker.Pick("Select Server Folder", Handle);
        if (p != null) tbServer.Text = p;
    }

    // ── Run robocopy (Sync tab) ───────────────────────────────────────────────
    void RunSync(string mode)
    {
        if (_isRunning)
        {
            MessageBox.Show("A sync or queue job is already running. Please wait.", "Busy",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string local  = tbLocal.Text.Trim();
        string server = tbServer.Text.Trim();

        if (string.IsNullOrEmpty(local) || string.IsNullOrEmpty(server))
        {
            MessageBox.Show("Please enter both Local and Server paths.", "Missing Paths",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string ts      = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string logDir  = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "log");
        if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
        string logFile = Path.Combine(logDir, "robocopy_log_" + ts + ".txt");
        currentLogFile = logFile;

        string args, label;
        bool   isFolder = rbFolder.Checked;

        if (isFolder)
        {
            string src = (mode != "server_to_local") ? local  : server;
            string dst = (mode != "server_to_local") ? server : local;

            if (mode == "move_to_server")
            {
                args  = Quote(src) + " " + Quote(dst) + " /E /MOVE /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
                label = "Local -> Server  +  Delete Local  [MOVE]";
            }
            else
            {
                args  = Quote(src) + " " + Quote(dst) + " /MIR /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
                label = (mode == "local_to_server") ? "Local -> Server  [MIRROR]" : "Server -> Local  [MIRROR]";
            }
        }
        else
        {
            string srcDir, srcFile, dstDir;
            if (mode == "local_to_server")
            {
                srcDir  = Path.GetDirectoryName(local);
                srcFile = Path.GetFileName(local);
                dstDir  = server;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
                label   = "File Copy: " + srcFile + "  ->  Server";
            }
            else if (mode == "server_to_local")
            {
                srcDir  = Path.GetDirectoryName(server);
                srcFile = Path.GetFileName(server);
                dstDir  = local;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
                label   = "File Copy: " + srcFile + "  ->  Local";
            }
            else
            {
                srcDir  = Path.GetDirectoryName(local);
                srcFile = Path.GetFileName(local);
                dstDir  = server;
                args    = Quote(srcDir) + " " + Quote(dstDir) + " " + Quote(srcFile) + " /MOV /R:3 /W:5 /TEE /LOG:" + Quote(logFile);
                label   = "File Move: " + srcFile + "  ->  Server  +  Delete Local";
            }
        }

        Config.Save(local, server);
        _isRunning = true;
        SetSyncButtonsEnabled(false);
        SetQueueButtonsEnabled(false);
        btnOpenLog.Enabled       = false;
        btnStop.Enabled          = true;
        lblCurrentFile.Text      = "";
        lblCurrentFile.ForeColor = C.Muted;
        pbFile.Value             = 0;
        lblPercent.Text          = "0%";
        _latestPct               = 0;
        _latestFile              = "";
        System.Threading.Interlocked.Exchange(ref _latestFileSize, 0);
        _fileChanged             = false;
        _progressTimer.Start();
        SetStatus("Running...", C.Yellow);

        string divider = new string('-', 60);
        AppendLog("\n" + divider + "\n", C.Muted);
        AppendLog("  " + label + "\n", C.Accent);
        AppendLog("  " + DateTime.Now.ToString("yyyy-MM-dd  HH:mm:ss") + "\n", C.Muted);
        AppendLog(divider + "\n\n", C.Muted);

        string capturedArgs = args;
        string capturedLog  = logFile;
        ThreadPool.QueueUserWorkItem(delegate { Worker(capturedArgs, capturedLog); });
    }

    static string Quote(string s) { return "\"" + s.TrimEnd('\\') + "\""; }

    // ── Worker thread (Sync tab) ──────────────────────────────────────────────
    void Worker(string args, string logFile)
    {
        try
        {
            var psi = new ProcessStartInfo("robocopy", args) {
                UseShellExecute        = false,
                RedirectStandardOutput = true,
                RedirectStandardError  = true,
                CreateNoWindow         = true,
            };

            using (var proc = Process.Start(psi))
            {
                Invoke((Action)delegate { activeProcess = proc; });

                proc.OutputDataReceived += delegate(object sender, DataReceivedEventArgs e)
                {
                    if (e.Data == null) return;
                    string raw = e.Data;
                    string t   = raw.TrimEnd('\r').Trim();
                    if (t.Length == 0) return;

                    if (t.EndsWith("%"))
                    {
                        string digits = t.TrimEnd('%').Trim();
                        int pct;
                        if (int.TryParse(digits, out pct) && pct >= 0 && pct <= 100)
                            _latestPct = pct;
                        return;
                    }

                    string filename, sizeRaw;
                    ExtractNameAndSize(raw, out filename, out sizeRaw);
                    if (filename != null)
                    {
                        _latestPct      = 0;
                        _latestFile     = filename;
                        System.Threading.Interlocked.Exchange(ref _latestFileSize, ParseRoboSize(sizeRaw));
                        _fileChanged    = true;
                    }

                    string captured = raw;
                    BeginInvoke((Action)delegate { ProcessRobocopyLine(captured); });
                };

                proc.BeginOutputReadLine();
                proc.WaitForExit();
                int rc = proc.ExitCode;
                Invoke((Action)delegate { Finish(rc, logFile); });
            }
        }
        catch (Exception ex)
        {
            string msg = ex.Message;
            Invoke((Action)delegate {
                AppendLog("\nError: " + msg + "\n", C.Red);
                Finish(-1, logFile);
            });
        }
    }

    // ── Parse and colour-code each robocopy line (Sync) ───────────────────────
    void ProcessRobocopyLine(string line)
    {
        string t = line.Trim();
        if (t.Length == 0) { AppendLog("\n", C.Text); return; }

        if (t.EndsWith("%") && !t.Contains(" ")) return;

        Color  color    = C.Text;

        if (t.IndexOf("New File",  StringComparison.OrdinalIgnoreCase) >= 0 ||
            t.IndexOf("New Dir",   StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Green;
        else if (t.IndexOf("Newer",  StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("Older",  StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("Changed",StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Yellow;
        else if (t.IndexOf("*EXTRA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("*Skipping",StringComparison.OrdinalIgnoreCase)>= 0)
            color = C.Muted;
        else if (t.IndexOf("FAILED",StringComparison.OrdinalIgnoreCase) >= 0 ||
                 t.IndexOf("ERROR", StringComparison.OrdinalIgnoreCase) >= 0)
            color = C.Red;
        else if (t.StartsWith("Dirs :") || t.StartsWith("Files :") ||
                 t.StartsWith("Bytes :") || t.StartsWith("Times :"))
            color = C.Accent;
        else if (t.StartsWith("---") || t.StartsWith("ROBOCOPY") ||
                 t.StartsWith("Started") || t.StartsWith("Ended") ||
                 t.StartsWith("Source") || t.StartsWith("Dest") ||
                 t.StartsWith("Files :") || t.StartsWith("Options"))
            color = C.Muted;

        AppendLog(line + "\n", color);
    }

    static string ExtractName(string line)
    {
        string name, size;
        ExtractNameAndSize(line, out name, out size);
        return name;
    }

    static void ExtractNameAndSize(string line, out string name, out string sizeRaw)
    {
        name = null; sizeRaw = null;
        string[] parts = line.Split('\t');
        int found = 0;
        for (int i = parts.Length - 1; i >= 0; i--)
        {
            string s = parts[i].Trim();
            if (s.Length == 0) continue;
            if (found == 0) { name    = s; found++; }
            else            { sizeRaw = s; return;  }
        }
    }

    static long ParseRoboSize(string raw)
    {
        if (raw == null) return 0;
        raw = raw.Trim().ToLower().Replace(",", "");
        double mult = 1;
        if      (raw.EndsWith("g")) { mult = 1073741824L; raw = raw.Substring(0, raw.Length - 1).Trim(); }
        else if (raw.EndsWith("m")) { mult = 1048576;     raw = raw.Substring(0, raw.Length - 1).Trim(); }
        else if (raw.EndsWith("k")) { mult = 1024;        raw = raw.Substring(0, raw.Length - 1).Trim(); }
        double val;
        if (double.TryParse(raw.Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out val))
            return (long)(val * mult);
        return 0;
    }

    static string FormatBytes(long bytes)
    {
        if (bytes >= 1073741824L) return string.Format("{0:F2} GB", bytes / 1073741824.0);
        if (bytes >= 1048576)     return string.Format("{0:F1} MB", bytes / 1048576.0);
        if (bytes >= 1024)        return string.Format("{0:F0} KB", bytes / 1024.0);
        return bytes + " B";
    }

    // ── Finish (Sync tab) ─────────────────────────────────────────────────────
    void Finish(int rc, string logFile)
    {
        _progressTimer.Stop();
        activeProcess       = null;
        _isRunning          = false;
        btnStop.Enabled     = false;
        lblCurrentFile.Text = "";
        pbFile.Value        = rc <= 7 ? 100 : 0;
        lblPercent.Text     = rc <= 7 ? "100%" : "";
        AppendLog("\n");

        if (rc <= 1)
        {
            AppendLog("  Sync completed successfully.\n", C.Green);
            SetStatus("Completed successfully", C.Green);
        }
        else if (rc <= 3)
        {
            AppendLog("  Sync completed - some files copied.\n", C.Green);
            SetStatus("Completed - files copied", C.Green);
        }
        else if (rc <= 7)
        {
            AppendLog("  Completed with warnings.\n", C.Yellow);
            SetStatus("Completed with warnings", C.Yellow);
        }
        else
        {
            AppendLog("  Sync errors (exit code " + rc + "). Check the log.\n", C.Red);
            SetStatus("Error (code " + rc + ")", C.Red);
        }
        AppendLog("\n  Log: " + logFile + "\n", C.Muted);
        SetSyncButtonsEnabled(true);
        SetQueueButtonsEnabled(true);
        btnOpenLog.Enabled = true;
    }

    // ── Log helpers (Sync) ────────────────────────────────────────────────────
    void AppendLog(string text, Color? color = null)
    {
        rtbLog.SelectionStart  = rtbLog.TextLength;
        rtbLog.SelectionLength = 0;
        rtbLog.SelectionColor  = color.HasValue ? color.Value : C.Text;
        rtbLog.AppendText(text);
        rtbLog.ScrollToCaret();
    }

    void SetStatus(string text, Color color)
    {
        lblStatus.Text      = text;
        lblStatus.ForeColor = color;
    }

    void SetSyncButtonsEnabled(bool enabled)
    {
        btnL2S.Enabled  = enabled;
        btnS2L.Enabled  = enabled;
        btnMove.Enabled = enabled;
    }

    void StopSync()
    {
        Process proc = activeProcess;
        if (proc == null || proc.HasExited) return;
        try { proc.Kill(); } catch { }

        _progressTimer.Stop();
        _isRunning      = false;
        btnStop.Enabled = false;
        pbFile.Value    = 0;
        lblPercent.Text = "";

        string divider = new string('-', 60);
        AppendLog("\n" + divider + "\n", C.Muted);
        AppendLog("  STOPPED by user.\n", C.Red);
        AppendLog(divider + "\n", C.Muted);

        SetStatus("Stopped", C.Red);
        SetSyncButtonsEnabled(true);
        SetQueueButtonsEnabled(true);
        btnOpenLog.Enabled   = true;
        lblCurrentFile.Text  = "";
        activeProcess        = null;
    }

    void OpenLog()
    {
        if (!string.IsNullOrEmpty(currentLogFile) && File.Exists(currentLogFile))
            Process.Start(currentLogFile);
    }
}

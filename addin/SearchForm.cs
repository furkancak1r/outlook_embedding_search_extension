// HatırMail window (UI thread only; Outlook objects are touched only here) and the worker session.
// Tabs: Ara (search), İndeks (index state, folders, full indexing started by the user), Ayarlar (settings.json).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace HatirMail
{
    public class SearchForm : Form
    {
        static readonly CultureInfo TR = CultureInfo.GetCultureInfo("tr-TR");
        static readonly Color Accent = Color.FromArgb(0, 114, 198), AccentDark = Color.FromArgb(0, 90, 158),
                              MarkColor = Color.FromArgb(255, 226, 110), Warn = Color.FromArgb(166, 77, 0),
                              Good = Color.FromArgb(16, 124, 16);
        static readonly bool Plain = SystemInformation.HighContrast; // high-contrast themes keep the system colors
        public static float TestScale; // test hook: > 0 simulates a high-DPI screen (sizes and fonts x TestScale) on a 96-DPI screen

        readonly IMailSource source;
        readonly Settings settings;
        readonly float k, pt;
        readonly Font titleFont, boldFont;
        readonly Panel root = new Panel { Dock = DockStyle.Fill };
        readonly TabControl tabs = new TabControl();
        readonly TabPage tabSearch = new TabPage("Ara"), tabIndex = new TabPage("İndeks"), tabSettings = new TabPage("Ayarlar");
        // Ara
        readonly TextBox query = new TextBox();
        readonly Button searchBtn, stopBtn, openBtn, historyBtn;
        readonly ContextMenuStrip historyMenu = new ContextMenuStrip { ShowItemToolTips = true };
        SearchHistory history; // recent and saved searches (data\searches.json)
        readonly Label scopeLine = new Label();
        readonly ListView list = new ListView();
        readonly RichTextBox preview = new RichTextBox();
        // İndeks
        readonly ComboBox storeBox = new ComboBox();
        readonly ListView folderList = new ListView();
        readonly ProgressBar barRead = new ProgressBar(), barSem = new ProgressBar(), barAtt = new ProgressBar();
        readonly Label pctRead = new Label(), pctSem = new Label(), pctAtt = new Label(), phase = new Label();
        readonly Button indexBtn, indexStopBtn, refreshBtn;
        // Ayarlar
        readonly NumericUpDown numThreads = new NumericUpDown(), numMemory = new NumericUpDown(), numBudget = new NumericUpDown(), numKeep = new NumericUpDown();
        readonly CheckBox chkAtt = new CheckBox();
        readonly Button saveBtn, defaultsBtn, deleteBtn;
        // Status bar
        readonly Label status = new Label(), summary = new Label();
        readonly ProgressBar progress = new ProgressBar();

        List<FolderInfo> folders = new List<FolderInfo>();
        readonly Dictionary<string, FolderCov> cov = new Dictionary<string, FolderCov>();
        WorkerSession session;
        WorkerProcess warm; // kept worker between sessions (window open, keep-alive on): the model stays loaded
        readonly System.Windows.Forms.Timer keepTimer = new System.Windows.Forms.Timer();
        Action queued;      // work that waits for the end of the current session (one worker at a time)
        string lastStoreId;
        bool loadingFolders, allowClose, closing, modelSeen, summaryQueued;
        public int KeepAliveMsOverride;    // test hook: keep-alive time in ms instead of the setting
        public event Action<string> ModelStateChanged; // test hook
        readonly Dictionary<Font, Font> markFonts = new Dictionary<Font, Font>(); // bold variants for the preview marks
        public double WorkerPeakMb; // test hook: peak RSS reported by the worker
        public readonly List<string> EventNames = new List<string>(); // test hook: event names, never mail text
        public event Action<string, string> SessionFinished; // test hook: (mode, reason)
        public event Action<string, int> WorkerStarted;      // test hook: (mode, worker process id)

        public SearchForm(IMailSource source, Settings settings)
        {
            this.source = source;
            this.settings = settings;
            // Sizes below are 96-DPI pixels; S() converts them to the screen DPI. Outlook is DPI-aware, so Windows does
            // not stretch this window, and fonts (in points) follow the DPI already. No WinForms auto-scaling: in a 2x
            // test it scaled a label in a nested panel 4x. ponytail: one DPI per window (system DPI); a move to a
            // monitor with another DPI keeps the size. Upgrade path: handle WM_DPICHANGED.
            k = TestScale > 0 ? TestScale : ScreenDpi() / 96F;
            pt = TestScale > 0 ? TestScale : 1F;
            Text = "HatırMail";
            Width = S(1040); Height = S(720); MinimumSize = new Size(S(820), S(560));
            Font = new Font("Segoe UI", 9f * pt);
            titleFont = new Font("Segoe UI Semibold", 11.5f * pt);
            boldFont = new Font("Segoe UI", 9f * pt, Plain ? FontStyle.Bold | FontStyle.Underline : FontStyle.Bold);
            StartPosition = FormStartPosition.CenterScreen;
            int line = TextRenderer.MeasureText("Ag", Font).Height;

            searchBtn = MakeButton("&Ara", true); stopBtn = MakeButton("D&urdur", false); openBtn = MakeButton("Outlook'ta A&ç", false);
            historyBtn = MakeButton("Ara&malar ▼", false);
            indexBtn = MakeButton("İndekslemeyi &başlat", true); indexStopBtn = MakeButton("&Durdur", false);
            refreshBtn = MakeButton("Durumu &yenile", false);
            saveBtn = MakeButton("&Kaydet", true); defaultsBtn = MakeButton("&Varsayılan değerler", false);
            deleteBtn = MakeButton("İndeksi &sil", false);

            // Header band
            var title = new Label { Text = "HatırMail", AutoSize = true, Dock = DockStyle.Top, Font = new Font("Segoe UI Semibold", 15f * pt), UseCompatibleTextRendering = false,
                                    ForeColor = Plain ? SystemColors.ControlText : Color.White };
            var subtitle = new Label { AutoSize = true, Dock = DockStyle.Top, UseCompatibleTextRendering = false, ForeColor = Plain ? SystemColors.ControlText : Color.FromArgb(222, 236, 249),
                                       Text = "Hatırladığınız birkaç ayrıntıyı yazın; e-postayı bulun ve Outlook'ta açın. Her şey bu bilgisayarda çalışır." };
            var header = new Panel { Dock = DockStyle.Top, BackColor = Plain ? SystemColors.Control : Accent, Padding = new Padding(S(16), S(8), S(16), S(10)) };
            header.Height = header.Padding.Vertical + title.PreferredHeight + subtitle.PreferredHeight + S(2);
            header.Controls.Add(subtitle); header.Controls.Add(title);

            // Status bar: the current work on the left, the index summary on the right
            var bar = new Panel { Dock = DockStyle.Bottom, Height = line + S(12), Padding = new Padding(S(12), S(4), S(12), S(4)) };
            status.Dock = DockStyle.Fill; status.AutoEllipsis = true; status.TextAlign = ContentAlignment.MiddleLeft;
            summary.Dock = DockStyle.Right; summary.AutoSize = true; summary.TextAlign = ContentAlignment.MiddleRight;
            summary.Padding = new Padding(S(12), 0, 0, 0);
            progress.Dock = DockStyle.Right; progress.Width = S(160); progress.Visible = false;
            bar.Controls.Add(status); bar.Controls.Add(progress); bar.Controls.Add(summary);
            var rule = new Panel { Dock = DockStyle.Bottom, Height = Math.Max(1, S(1)), BackColor = SystemColors.ControlDark };

            tabs.Dock = DockStyle.Fill; tabs.Padding = new Point(S(14), S(5));
            tabs.TabPages.Add(tabSearch); tabs.TabPages.Add(tabIndex); tabs.TabPages.Add(tabSettings);
            foreach (TabPage p in tabs.TabPages) { p.Padding = new Padding(S(12)); p.UseVisualStyleBackColor = true; }

            // ---- Ara: query row, hint, scope line, results, preview
            var qrow = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 5, RowCount = 1, AutoSize = true };
            qrow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 4; i++) qrow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            query.Font = new Font("Segoe UI", 12f * pt); query.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            query.Margin = new Padding(0, S(3), S(6), S(3)); query.AccessibleName = "Arama metni";
            query.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; StartSearch(); } };
            // Suggestions while typing: the Windows autocomplete list (prefix match) with recent and saved searches.
            history = SearchHistory.Load(settings.Data);
            query.AutoCompleteMode = AutoCompleteMode.Suggest; query.AutoCompleteSource = AutoCompleteSource.CustomSource;
            RefreshSuggestions();
            historyBtn.AccessibleName = "Son ve kayıtlı aramalar"; historyBtn.AccessibleRole = AccessibleRole.ButtonMenu;
            historyBtn.Click += (s, e) => { BuildHistoryMenu(); historyMenu.Show(historyBtn, new Point(0, historyBtn.Height)); };
            historyMenu.Font = Font; // the window font (DPI test scale included), not the system menu font
            Disposed += (s, e) => historyMenu.Dispose(); // not a child control
            qrow.Controls.Add(query, 0, 0); qrow.Controls.Add(historyBtn, 1, 0); qrow.Controls.Add(searchBtn, 2, 0);
            qrow.Controls.Add(stopBtn, 3, 0); qrow.Controls.Add(openBtn, 4, 0);
            var hint = Wrap(new Label { ForeColor = SystemColors.GrayText, Padding = new Padding(0, S(4), 0, 0),
                                        Text = "İpucu: kişi, konu, yaklaşık tarih ya da ekteki bir sözcük yazın. Örnek: geçen ay Ahmet'ten gelen kargo faturası" });
            Wrap(scopeLine); scopeLine.Padding = new Padding(0, S(4), 0, S(8));
            scopeLine.Cursor = Cursors.Hand; scopeLine.Click += (s, e) => tabs.SelectedTab = tabIndex;
            list.Dock = DockStyle.Fill; list.View = View.Details; list.FullRowSelect = true; list.MultiSelect = false;
            list.HideSelection = false; list.AccessibleName = "Sonuçlar";
            // "Neden" fits its longest form ("kelime 1/2, anlam 0.61, ekte, tarih dışı"); Konu takes the rest.
            list.Columns.Add("Konu", S(300)); list.Columns.Add("Kişi", S(140)); list.Columns.Add("Tarih", S(115));
            list.Columns.Add("Güven", S(65)); list.Columns.Add("Neden", S(265));
            list.Resize += (s, e) => FitColumns(list);
            list.SelectedIndexChanged += (s, e) => ShowPreview();
            list.DoubleClick += (s, e) => OpenSelected();
            list.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) OpenSelected(); };
            // Plain text only (never RTF from a mail), read-only, no links.
            preview.Dock = DockStyle.Bottom; preview.Height = S(190); preview.ReadOnly = true; preview.DetectUrls = false;
            preview.BorderStyle = BorderStyle.FixedSingle; preview.BackColor = SystemColors.Window; preview.AccessibleName = "Önizleme";
            preview.ScrollBars = RichTextBoxScrollBars.Vertical;
            qrow.TabIndex = 0; list.TabIndex = 1; preview.TabIndex = 2;
            tabSearch.Controls.Add(list); tabSearch.Controls.Add(new Splitter { Dock = DockStyle.Bottom, Height = S(6), TabStop = false });
            tabSearch.Controls.Add(preview); tabSearch.Controls.Add(scopeLine); tabSearch.Controls.Add(hint); tabSearch.Controls.Add(qrow);

            // ---- İndeks: store, bars, live phase, start/stop, folders with their state
            var srow = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, RowCount = 1, AutoSize = true };
            srow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); srow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            srow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            storeBox.DropDownStyle = ComboBoxStyle.DropDownList; storeBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            storeBox.AccessibleName = "Posta kutusu"; storeBox.SelectedIndexChanged += (s, e) => LoadFolders();
            srow.Controls.Add(new Label { Text = "Posta kutusu:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, S(6), 0) }, 0, 0);
            srow.Controls.Add(storeBox, 1, 0); srow.Controls.Add(refreshBtn, 2, 0);
            var grid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, RowCount = 3, AutoSize = true, Padding = new Padding(0, S(8), 0, S(4)) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, S(250)));
            BarRow(grid, 0, "Okunan postalar", barRead, pctRead);
            BarRow(grid, 1, "Anlam indeksi", barSem, pctSem);
            BarRow(grid, 2, "Ek dosyaları", barAtt, pctAtt);
            Wrap(phase); phase.Padding = new Padding(0, S(4), 0, S(6)); phase.Font = boldFont; phase.Text = "İndeksleme çalışmıyor.";
            var brow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false };
            brow.Controls.Add(indexBtn); brow.Controls.Add(indexStopBtn);
            var note = Wrap(new Label { ForeColor = SystemColors.GrayText, Padding = new Padding(0, S(2), 0, S(8)),
                                        Text = "Tam indeksleme uzun sürebilir; bu sırada Outlook'u kullanabilirsiniz. Durdurursanız yapılan iş kaydedilir ve sonra kaldığı yerden sürer. Outlook kapanınca indeksleme de durur." });
            var flabel = Wrap(new Label { Padding = new Padding(0, 0, 0, S(4)), Text = "Klasörler (işaretli klasörler aranır ve indekslenir):" });
            folderList.Dock = DockStyle.Fill; folderList.View = View.Details; folderList.CheckBoxes = true; folderList.FullRowSelect = true;
            folderList.HideSelection = false; folderList.AccessibleName = "Klasörler";
            folderList.Columns.Add("Klasör", S(330));
            foreach (var c in new[] { "Posta", "Okunan", "Anlam", "Ekler" }) folderList.Columns.Add(c, S(c == "Ekler" ? 110 : 80), HorizontalAlignment.Right);
            folderList.Columns.Add("Son tarama", S(130));
            folderList.Resize += (s, e) => FitColumns(folderList);
            // ItemChecked also comes while Windows builds the native list (first show of the İndeks tab, with visual
            // styles): some items are not there yet (Outlook, 16 Oct: NullReferenceException). The summary runs once,
            // after that message. A running search or index run keeps the folder scope of its start.
            folderList.ItemChecked += (s, e) => { if (!loadingFolders) QueueSummary(); };
            srow.TabIndex = 0; brow.TabIndex = 1; folderList.TabIndex = 2;
            tabIndex.Controls.Add(folderList); tabIndex.Controls.Add(flabel); tabIndex.Controls.Add(note); tabIndex.Controls.Add(brow);
            tabIndex.Controls.Add(phase); tabIndex.Controls.Add(grid); tabIndex.Controls.Add(srow);

            // ---- Ayarlar: saved to settings.json
            var sgrid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, AutoSize = true };
            sgrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); sgrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            numThreads.Minimum = 1; numThreads.Maximum = 16;
            numMemory.Minimum = 2048; numMemory.Maximum = 32768; numMemory.Increment = 512;
            numBudget.Minimum = 0; numBudget.Maximum = 600; numBudget.Increment = 5;
            numKeep.Minimum = 0; numKeep.Maximum = 120;
            SettingRow(sgrid, 0, "İşlemci çekirdeği:", numThreads, string.Format(TR,
                "1–16. Bu bilgisayarda {0} mantıksal çekirdek var. Çok çekirdek indekslemeyi hızlandırır ama Outlook'u yavaşlatabilir.",
                Environment.ProcessorCount));
            SettingRow(sgrid, 1, "Bellek sınırı (MB):", numMemory, "Arama işlemi bu sınırı aşarsa kapatılır. Yapay zekâ modeliyle indeksleme en çok yaklaşık 1.300 MB kullanır.");
            SettingRow(sgrid, 2, "Aramada indeksleme (sn):", numBudget,
                "Ara'ya bastığınızda yeni postaları indekslemek için en çok bu kadar beklenir; kalan iş sonraki aramaya ya da İndeks sekmesine kalır. 0: aramada indeksleme yok.");
            SettingRow(sgrid, 3, "Yapay zekâyı hazır tut (dk):", numKeep,
                "Pencere açıkken model son işten sonra bu kadar dakika bellekte kalır; sonraki arama beklemeden çalışır. Pencere kapanınca model de kapanır. 0: her aramada yeniden yüklenir.");
            chkAtt.Text = "Ek dosyalarının içinde de ara (Excel, PDF, Word, PowerPoint, resimler ve taranmış belgeler için OCR, zip…)";
            chkAtt.AutoSize = true; chkAtt.Margin = new Padding(0, S(10), 0, S(10));
            sgrid.Controls.Add(chkAtt, 0, 8); sgrid.SetColumnSpan(chkAtt, 2);
            var sbtn = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            deleteBtn.Margin = new Padding(S(24), deleteBtn.Margin.Top, deleteBtn.Margin.Right, deleteBtn.Margin.Bottom); // apart from Kaydet
            sbtn.Controls.Add(saveBtn); sbtn.Controls.Add(defaultsBtn); sbtn.Controls.Add(deleteBtn);
            sgrid.Controls.Add(sbtn, 0, 9); sgrid.SetColumnSpan(sbtn, 2);
            var info = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, S(14), 0, 0), Anchor = AnchorStyles.Left | AnchorStyles.Right,
                                   Text = "Gizlilik: Arama, ek okuma ve yapay zekâ yalnız bu bilgisayarda çalışır. E-posta ve arama metni hiçbir yere gönderilmez ve günlüğe yazılmaz.\n" +
                                          "İndeks dosyaları: " + Path.Combine(settings.Data, "index") + "\nAyarlar dosyası: " + settings.ConfigFile };
            sgrid.Controls.Add(info, 0, 10); sgrid.SetColumnSpan(info, 2);
            tabSettings.Controls.Add(sgrid);

            root.Controls.Add(tabs); root.Controls.Add(rule); root.Controls.Add(bar); root.Controls.Add(header);
            Controls.Add(root);
            searchBtn.Click += (s, e) => StartSearch();
            stopBtn.Click += (s, e) => StopSession();
            indexStopBtn.Click += (s, e) => StopSession();
            openBtn.Click += (s, e) => OpenSelected();
            indexBtn.Click += (s, e) => StartIndex();
            refreshBtn.Click += (s, e) => StartStatus();
            saveBtn.Click += (s, e) => SaveSettings();
            deleteBtn.Click += (s, e) => ConfirmDeleteIndex();
            defaultsBtn.Click += (s, e) => { ShowSettings(new Settings()); SetStatus("Varsayılan değerler yüklendi. Kaydetmek için Kaydet'e basın."); };
            FormClosing += OnClosing;
            Load += (s, e) =>
            {
                // A scaled window can be larger than a small high-DPI screen (1366x768 at 125%): fit it to the work area.
                var wa = Screen.FromControl(this).WorkingArea;
                MinimumSize = new Size(Math.Min(MinimumSize.Width, wa.Width), Math.Min(MinimumSize.Height, wa.Height));
                Size = new Size(Math.Min(Width, wa.Width), Math.Min(Height, wa.Height));
                ShowSettings(settings);
                LoadStores(); // also starts the index status check of the first store (a short worker run, no model)
                ShowPreview();
                FitColumns(list); FitColumns(folderList);
                ActiveControl = query;
            };
            // GDI text (as TextRenderer measures it): Outlook does not call SetCompatibleTextRenderingDefault(false).
            foreach (var c in AllControls(this))
            {
                if (c is Label) ((Label)c).UseCompatibleTextRendering = false;
                if (c is ButtonBase) ((ButtonBase)c).UseCompatibleTextRendering = false;
            }
            SetStatus("Hazır. Hatırladıklarınızı yazıp Ara'ya basın. Yapay zekâ modeli yalnız bu pencere açıkken çalışır.");
            keepTimer.Tick += (s, e) => { keepTimer.Stop(); CloseWarm("keep-alive"); };
            UpdateButtons();
        }

        public static IEnumerable<Control> AllControls(Control c)
        {
            foreach (Control x in c.Controls) { yield return x; foreach (var y in AllControls(x)) yield return y; }
        }

        // A docked label that wraps its text at its width (an AutoSize label stays on one line and is cut).
        Label Wrap(Label l)
        {
            l.AutoSize = false; l.Dock = DockStyle.Top;
            EventHandler fit = (s, e) =>
            {
                int h = WrapHeight(l);
                if (l.Height != h) l.Height = h;
            };
            l.Resize += fit; l.TextChanged += fit; l.FontChanged += fit; l.PaddingChanged += fit;
            return l;
        }

        public static int WrapHeight(Label l)
        {
            int w = Math.Max(1, l.ClientSize.Width - l.Padding.Horizontal);
            return TextRenderer.MeasureText(l.Text.Length > 0 ? l.Text : " ", l.Font, new Size(w, int.MaxValue),
                                            TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height + l.Padding.Vertical + 1;
        }

        int S(int px) { return (int)Math.Round(px * k); }
        static float ScreenDpi() { using (var g = Graphics.FromHwnd(IntPtr.Zero)) return g.DpiX; }

        Button MakeButton(string text, bool primary)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, S(2), S(6), S(2)),
                                 Padding = new Padding(S(10), S(3), S(10), S(3)), UseVisualStyleBackColor = !primary || Plain };
            if (primary && !Plain)
            {
                b.FlatStyle = FlatStyle.Flat; b.BackColor = Accent; b.ForeColor = Color.White; b.Font = boldFont ?? Font;
                b.FlatAppearance.BorderColor = AccentDark; b.FlatAppearance.MouseOverBackColor = AccentDark;
                // A disabled flat button keeps its colors: show the state by color (EnabledChanged).
                b.EnabledChanged += (s, e) => { b.BackColor = b.Enabled ? Accent : SystemColors.ControlLight; b.ForeColor = b.Enabled ? Color.White : SystemColors.GrayText; };
            }
            return b;
        }

        void BarRow(TableLayoutPanel g, int row, string name, ProgressBar bar, Label pct)
        {
            var l = new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, S(4), S(10), S(4)) };
            bar.Anchor = AnchorStyles.Left | AnchorStyles.Right; bar.Height = S(18); bar.Margin = new Padding(0, S(4), S(10), S(4));
            bar.Maximum = 1000; bar.AccessibleName = name;
            pct.AutoSize = true; pct.Anchor = AnchorStyles.Left; pct.Text = "–";
            g.Controls.Add(l, 0, row); g.Controls.Add(bar, 1, row); g.Controls.Add(pct, 2, row);
        }

        void SettingRow(TableLayoutPanel g, int row, string name, NumericUpDown num, string help)
        {
            num.Width = S(110); num.Margin = new Padding(0, S(6), 0, 0); num.AccessibleName = name.TrimEnd(':');
            num.TextAlign = HorizontalAlignment.Right; num.ThousandsSeparator = false;
            g.Controls.Add(new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, S(8), S(12), 0) }, 0, row * 2);
            g.Controls.Add(num, 1, row * 2);
            var h = new Label { Text = help, AutoSize = true, ForeColor = SystemColors.GrayText, Anchor = AnchorStyles.Left | AnchorStyles.Right,
                                Margin = new Padding(0, S(2), 0, S(6)) }; // anchored left and right: wraps at the cell width
            g.Controls.Add(h, 1, row * 2 + 1);
        }

        // Columns fill the list width; the first column takes the change (no column cut off at the right edge).
        static void FitColumns(ListView lv)
        {
            if (lv.Columns.Count == 0 || lv.ClientSize.Width <= 0) return;
            int rest = 0;
            for (int i = 1; i < lv.Columns.Count; i++) rest += lv.Columns[i].Width;
            lv.Columns[0].Width = Math.Max(lv.Columns[1].Width, lv.ClientSize.Width - rest - 2);
        }

        // ------------------------------------------------------------------ settings
        void ShowSettings(Settings s)
        {
            numThreads.Value = Math.Max(numThreads.Minimum, Math.Min(numThreads.Maximum, s.Threads));
            numMemory.Value = Math.Max(numMemory.Minimum, Math.Min(numMemory.Maximum, s.MemoryLimitMb));
            numBudget.Value = Math.Max(numBudget.Minimum, Math.Min(numBudget.Maximum, s.SearchIndexBudgetS));
            numKeep.Value = Math.Max(numKeep.Minimum, Math.Min(numKeep.Maximum, s.KeepAliveMin));
            chkAtt.Checked = s.Attachments;
        }

        void SaveSettings()
        {
            var old = new[] { settings.Threads, settings.MemoryLimitMb, settings.SearchIndexBudgetS, settings.KeepAliveMin };
            bool oldAtt = settings.Attachments;
            settings.Threads = (int)numThreads.Value; settings.MemoryLimitMb = (int)numMemory.Value;
            settings.SearchIndexBudgetS = (int)numBudget.Value; settings.Attachments = chkAtt.Checked; settings.KeepAliveMin = (int)numKeep.Value;
            try
            {
                settings.Save();
                SetStatus(session != null ? "Ayarlar kaydedildi. Yeni değerler bir sonraki işte geçerli olur." : "Ayarlar kaydedildi.");
                // A kept worker has the old threads and memory limit (set at its start): it closes; keep-alive 0 closes it too.
                if (old[0] != settings.Threads || old[1] != settings.MemoryLimitMb || settings.KeepAliveMin == 0) CloseWarm("settings");
                else if (old[3] != settings.KeepAliveMin && warm != null) ArmKeep();
                if (oldAtt != settings.Attachments) StartStatus(); // attachment counts change
            }
            catch (Exception ex)
            {
                settings.Threads = old[0]; settings.MemoryLimitMb = old[1]; settings.SearchIndexBudgetS = old[2]; settings.KeepAliveMin = old[3];
                settings.Attachments = oldAtt;
                SetStatus("Ayarlar kaydedilemedi (" + ex.GetType().Name + "). Eski değerler geçerli.");
            }
        }

        void ConfirmDeleteIndex()
        {
            if (session != null) return;
            if (MessageBox.Show(this, "Bu bilgisayardaki arama indeksi silinsin mi?\n\n" +
                                "Silinecekler: tüm posta kutularının indeks dosyaları (posta metni, kelime ve anlam indeksi, eklerden okunan metin).\n" +
                                "Silinmeyecekler: Outlook'taki postalarınız ve ekleri, ayarlarınız, son ve kayıtlı aramalarınız.\n\n" +
                                "Sonra arama yapmak için İndeks sekmesinden indekslemeyi baştan başlatmanız gerekir. Büyük bir posta kutusunda bu saatler sürebilir.",
                                "HatırMail", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            DeleteIndex();
        }

        /// <summary>Delete data\index (one SQLite file per store: mails, keyword index, vectors, attachment text) and
        /// data\tmp-att. Outlook mail and settings.json stay. No worker may hold the files: the button is off during a
        /// session, and the kept worker closes first (its model has nothing to search without an index).
        /// ponytail: a file held by another program (antivirus, backup) stops the delete half way; the status check
        /// below shows what is left, and the user presses again.</summary>
        public void DeleteIndex()
        {
            if (session != null) { SetStatus("Önce çalışan aramayı ya da indekslemeyi durdurun."); return; }
            CloseWarm("index deleted");
            try
            {
                foreach (var d in new[] { "index", "tmp-att" })
                {
                    string p = Path.Combine(settings.Data, d);
                    if (Directory.Exists(p)) Directory.Delete(p, true);
                }
                SetStatus("İndeks silindi. Outlook'taki postalarınız duruyor. Yeniden aramak için İndeks sekmesinde indekslemeyi başlatın.");
            }
            catch (Exception ex)
            {
                SetStatus("İndeks silinemedi (" + ex.GetType().Name + "). Başka bir program dosyaları kullanıyor olabilir. Biraz sonra yeniden deneyin.");
            }
            list.Items.Clear(); ShowPreview();
            cov.Clear(); UpdateSummary(); phase.Text = "İndeksleme çalışmıyor.";
            StartStatus(); // the real state; a status check makes no index file for an unknown store
        }

        // ------------------------------------------------------------------ recent and saved searches (data\searches.json)
        /// <summary>A search starts (after its checks): it goes first in the recent list. A write error does not stop
        /// the search; the list then shows what the file has.</summary>
        public void RememberSearch(string q) { Change(h => h.AddRecent(q), null); }

        /// <summary>Save the text of the search box under name ("Bu aramayı kaydet…"). The same name replaces the query
        /// of that saved search. Returns false when nothing was saved.</summary>
        public bool SaveSearchAs(string name)
        {
            string q = query.Text.Trim();
            name = (name ?? "").Trim();
            if (q.Length == 0 || name.Length == 0) return false;
            bool updated = false;
            if (!Change(h => updated = h.SetSaved(name, q), "Arama kaydedilemedi")) return false;
            SetStatus((updated ? "Kayıtlı arama güncellendi: " : "Arama kaydedildi: ") + name + ". Aramalar düğmesinden tek tıklamayla çalıştırabilirsiniz.");
            return true;
        }

        public void ClearRecent()
        {
            if (Change(h => h.Recent.Clear(), "Son aramalar silinemedi")) SetStatus("Son aramalar silindi. Kayıtlı aramalar duruyor.");
        }

        void DeleteSaved(string name)
        {
            if (Change(h => h.Saved.RemoveAll(x => x.Key == name), "Kayıtlı arama silinemedi")) SetStatus("Kayıtlı arama silindi: " + name + ".");
        }

        public void RunSearch(string q) { query.Text = q; query.SelectionStart = q.Length; StartSearch(); }

        /// <summary>Change a copy of the lists and write it; the window takes the copy only after a good write, so it
        /// never shows a change that is not saved. A file that could not be read at the window open is read again
        /// first: a write over it would delete its searches. failText null = no status text (recent searches: a write
        /// error does not stop the search).</summary>
        bool Change(Action<SearchHistory> edit, string failText)
        {
            if (history.ReadFailed)
            {
                var again = SearchHistory.Load(settings.Data);
                if (again.ReadFailed) { if (failText != null) SetStatus(failText + " (aramalar dosyası okunamadı). Biraz sonra yeniden deneyin."); return false; }
                history = again;
            }
            var next = history.Clone();
            edit(next);
            try { next.Save(); }
            catch (Exception ex)
            {
                if (failText != null) SetStatus(failText + " (" + ex.GetType().Name + "). Biraz sonra yeniden deneyin.");
                return false;
            }
            history = next;
            RefreshSuggestions();
            return true;
        }

        void RefreshSuggestions()
        {
            var c = new AutoCompleteStringCollection();
            c.AddRange(history.Recent.Concat(history.Saved.Select(x => x.Value)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
            // ponytail: a new collection, set after the current message: a change of the source while its list is open
            // (Enter in the search box starts the search) can crash WinForms. Before the window exists: at once.
            if (IsHandleCreated) BeginInvoke((Action)(() => query.AutoCompleteCustomSource = c));
            else query.AutoCompleteCustomSource = c;
        }

        // Built at each open: the lists and the button states are current.
        void BuildHistoryMenu()
        {
            foreach (var old in historyMenu.Items.Cast<ToolStripItem>().ToList()) old.Dispose(); // also the submenu window
            historyMenu.Items.Clear();
            bool canRun = searchBtn.Enabled; // as Ara: off while a search runs
            Func<string, ToolStripMenuItem> head = t => new ToolStripMenuItem(t) { Enabled = false, Font = boldFont };
            Func<string, string, ToolStripMenuItem> run = (text, q) =>
            {
                var it = new ToolStripMenuItem(MenuText(text)) { ToolTipText = q, Enabled = canRun };
                it.Click += (s, e) => RunSearch(q);
                return it;
            };
            historyMenu.Items.Add(head("Kayıtlı aramalar"));
            if (history.Saved.Count == 0) historyMenu.Items.Add(new ToolStripMenuItem("(kayıtlı arama yok)") { Enabled = false });
            foreach (var kv in history.Saved) historyMenu.Items.Add(run(kv.Key, kv.Value));
            var save = new ToolStripMenuItem("Bu aramayı &kaydet…") { Enabled = query.Text.Trim().Length > 0 };
            save.Click += (s, e) => AskSaveSearch();
            var del = new ToolStripMenuItem("Kayıtlı aramayı &sil") { Enabled = history.Saved.Count > 0 };
            foreach (var kv in history.Saved)
            {
                string name = kv.Key;
                var it = new ToolStripMenuItem(MenuText(name)) { ToolTipText = kv.Value };
                it.Click += (s, e) => DeleteSaved(name);
                del.DropDownItems.Add(it);
            }
            historyMenu.Items.Add(save); historyMenu.Items.Add(del);
            historyMenu.Items.Add(new ToolStripSeparator());
            historyMenu.Items.Add(head("Son aramalar"));
            if (history.Recent.Count == 0) historyMenu.Items.Add(new ToolStripMenuItem("(son arama yok)") { Enabled = false });
            foreach (var q in history.Recent) historyMenu.Items.Add(run(q, q));
            var clear = new ToolStripMenuItem("Son aramaları &temizle") { Enabled = history.Recent.Count > 0 };
            clear.Click += (s, e) => ClearRecent();
            historyMenu.Items.Add(clear);
        }

        /// <summary>Menu text: at most 60 characters; "&" stays a character (not a shortcut letter).</summary>
        public static string MenuText(string t) { return (t.Length > 60 ? t.Substring(0, 59) + "…" : t).Replace("&", "&&"); }

        void AskSaveSearch()
        {
            string q = query.Text.Trim();
            if (q.Length == 0) return;
            // The name of a saved search with this query, else the start of the query.
            string name = history.Saved.Where(x => string.Equals(x.Value, q, StringComparison.OrdinalIgnoreCase)).Select(x => x.Key).FirstOrDefault()
                          ?? (q.Length > 40 ? q.Substring(0, 40).TrimEnd() : q);
            using (var d = new NameDialog(this, name))
                if (d.ShowDialog(this) == DialogResult.OK) SaveSearchAs(d.Box.Text);
        }

        /// <summary>The name of a saved search (modal). Kaydet is off while the name is empty.</summary>
        public sealed class NameDialog : Form
        {
            public readonly TextBox Box = new TextBox();
            public readonly TableLayoutPanel Body = new TableLayoutPanel { AutoSize = true, ColumnCount = 1 };

            public NameDialog(SearchForm o, string name)
            {
                Text = "Aramayı kaydet"; Font = o.Font; AutoScaleMode = AutoScaleMode.None; FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = MinimizeBox = false; ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent;
                var l = new Label { AutoSize = true, Text = "Aramanın adı (Aramalar menüsünde görünür):", Margin = new Padding(0, 0, 0, o.S(4)), UseCompatibleTextRendering = false };
                Box.Width = o.S(380); Box.MaxLength = 80; Box.Text = name; Box.AccessibleName = "Aramanın adı"; Box.Margin = new Padding(0, 0, 0, o.S(12));
                var ok = o.MakeButton("&Kaydet", true); var cancel = o.MakeButton("İptal", false);
                ok.DialogResult = DialogResult.OK; cancel.DialogResult = DialogResult.Cancel;
                ok.UseCompatibleTextRendering = cancel.UseCompatibleTextRendering = false;
                var b = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.Right, Margin = new Padding(0) };
                b.Controls.Add(ok); b.Controls.Add(cancel);
                Body.Controls.Add(l); Body.Controls.Add(Box); Body.Controls.Add(b);
                Body.Location = new Point(o.S(14), o.S(12));
                Controls.Add(Body);
                var ps = Body.GetPreferredSize(Size.Empty);
                ClientSize = new Size(ps.Width + o.S(28), ps.Height + o.S(22));
                AcceptButton = ok; CancelButton = cancel;
                Box.TextChanged += (s, e) => ok.Enabled = Box.Text.Trim().Length > 0;
                ok.Enabled = name.Trim().Length > 0;
            }
        }

        // ------------------------------------------------------------------ scope
        void LoadStores()
        {
            try
            {
                storeBox.Items.Clear();
                foreach (var kv in source.Stores()) storeBox.Items.Add(new StoreItem { Id = kv.Key, Name = kv.Value });
                // Start with a store that has an index (indexed before), else with the first one.
                int pick = storeBox.Items.Cast<StoreItem>().ToList().FindIndex(x => File.Exists(IndexFile(x.Id)));
                if (storeBox.Items.Count > 0) storeBox.SelectedIndex = Math.Max(0, pick);
            }
            catch (Exception ex) { SetStatus("Posta kutuları okunamadı: " + ex.GetType().Name); }
        }

        void LoadFolders()
        {
            var st = storeBox.SelectedItem as StoreItem;
            if (session != null && session.Mode == "status") { queued = null; session.Kill("store changed"); }
            folderList.Items.Clear(); cov.Clear();
            if (st == null) return;
            try
            {
                folders = source.Folders(st.Id);
                loadingFolders = true;
                folderList.BeginUpdate();
                foreach (var f in folders)
                {
                    string p = f.Path.TrimStart('\\');
                    bool skip = p.StartsWith("Silinmiş", StringComparison.OrdinalIgnoreCase) || p.StartsWith("Deleted", StringComparison.OrdinalIgnoreCase)
                                || p.StartsWith("Önemsiz", StringComparison.OrdinalIgnoreCase) // Turkish Outlook: "Önemsiz E-Posta"
                                || p.StartsWith("Gereksiz", StringComparison.OrdinalIgnoreCase) || p.StartsWith("Junk", StringComparison.OrdinalIgnoreCase)
                                || p.StartsWith("Giden", StringComparison.OrdinalIgnoreCase) || p.StartsWith("Outbox", StringComparison.OrdinalIgnoreCase)
                                || p.StartsWith("Taslak", StringComparison.OrdinalIgnoreCase) || p.StartsWith("Drafts", StringComparison.OrdinalIgnoreCase)
                                || p.StartsWith("Eşitleme", StringComparison.OrdinalIgnoreCase) || p.StartsWith("Sync Issues", StringComparison.OrdinalIgnoreCase);
                    var it = new ListViewItem(new[] { p, f.Count.ToString("N0", TR), "", "", "", "" }) { Checked = !skip, Tag = f };
                    folderList.Items.Add(it);
                }
                folderList.EndUpdate();
                FitColumns(folderList);
                loadingFolders = false;
                UpdateSummary();
                StartStatus();
            }
            catch (Exception ex) { folderList.EndUpdate(); loadingFolders = false; SetStatus("Klasörler okunamadı: " + ex.GetType().Name); }
        }

        // The worker's file name for a store (hm_worker.db_path): sha256 of the UTF-8 store id, first 24 hex characters.
        string IndexFile(string storeId)
        {
            using (var sha = SHA256.Create())
            {
                string h = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(storeId))).Replace("-", "").ToLowerInvariant();
                return Path.Combine(settings.Data, "index", h.Substring(0, 24) + ".sqlite");
            }
        }

        void QueueSummary()
        {
            if (summaryQueued || !IsHandleCreated) return;
            summaryQueued = true;
            BeginInvoke((Action)(() => { summaryQueued = false; UpdateSummary(); }));
        }

        List<FolderInfo> SelectedFolders()
        {
            return folderList.Items.Cast<ListViewItem>().Where(i => i.Checked).Select(i => (FolderInfo)i.Tag).ToList();
        }

        // ------------------------------------------------------------------ sessions (one worker at a time)
        bool CanStart(out StoreItem st, out List<FolderInfo> sel)
        {
            st = storeBox.SelectedItem as StoreItem;
            sel = SelectedFolders();
            if (st == null || sel.Count == 0) { SetStatus("İndeks sekmesinde en az bir klasör seçin."); return false; }
            if (!File.Exists(settings.Python) || !File.Exists(settings.Worker))
            {
                SetStatus("Arama bileşeni bulunamadı. Kurulum klasörü: " + settings.Home);
                return false;
            }
            return true;
        }

        public void StartSearch()
        {
            string q = query.Text.Trim();
            if (q.Length == 0) { SetStatus("Önce hatırladığınız birkaç ayrıntıyı yazın."); query.Focus(); return; }
            if (session != null)
            {
                if (session.Mode == "search") return;
                if (session.Mode == "index")
                {
                    // One worker at a time: the index run stops (its work is saved), then the search runs.
                    if (MessageBox.Show(this, "İndeksleme sürüyor. Arama için indeksleme durdurulsun mu?\n\nYapılan iş kaydedilir. " +
                                        "İndekslemeyi sonra İndeks sekmesinden kaldığı yerden sürdürebilirsiniz.", "HatırMail",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                    queued = () => { query.Text = q; StartSearch(); };
                    StopSession();
                    return;
                }
                queued = () => { query.Text = q; StartSearch(); }; // a status check (about a second): the search follows it
                return;
            }
            StoreItem st; List<FolderInfo> sel;
            if (!CanStart(out st, out sel)) return;
            RememberSearch(q);
            list.Items.Clear(); ShowPreview();
            EventNames.Clear(); modelSeen = false;
            lastStoreId = st.Id;
            Begin(new WorkerSession(this, source, settings, "search", st.Id, st.Name, sel, q));
            SetStatus(ModelState == "loading" ? "Arama başladı. Yapay zekâ modeli hazırlanıyor (pencere açılınca başladı)…"
                    : ModelState == "ready" ? "Arama başladı. Yapay zekâ hazır." : "Arama başladı…");
        }

        public void StartIndex()
        {
            if (session != null)
            {
                if (session.Mode == "status") queued = StartIndex; // after the status check
                return;
            }
            StoreItem st; List<FolderInfo> sel;
            if (!CanStart(out st, out sel)) return;
            EventNames.Clear(); modelSeen = false;
            Begin(new WorkerSession(this, source, settings, "index", st.Id, st.Name, sel, ""));
            phase.Text = "İndeksleme başladı: klasör listeleri okunuyor…";
            SetStatus("İndeksleme başladı. Bu pencereyi kapatabilir ya da arama yapmadan bekleyebilirsiniz.");
        }

        /// <summary>Index state of the selected folders: a short worker run (no scan, no mail reading). With keep-alive on,
        /// the worker then loads the model in the background when the store has a semantic index (preload), so that the
        /// first search does not wait for it. Runs only on a user action: window open, store change, "Durumu yenile",
        /// after saving settings.</summary>
        public void StartStatus()
        {
            if (session != null) return;
            var st = storeBox.SelectedItem as StoreItem;
            var all = folders;
            if (st == null || all.Count == 0 || !File.Exists(settings.Python) || !File.Exists(settings.Worker)) return;
            EventNames.Clear();
            Begin(new WorkerSession(this, source, settings, "status", st.Id, st.Name, all, "", preload: true));
        }

        int KeepMs { get { return KeepAliveMsOverride > 0 ? KeepAliveMsOverride : settings.KeepAliveMin * 60000; } }

        // The worker's own limit for the wait between sessions. The window closes the worker first (keepTimer); this only
        // ends a worker whose window hangs. ponytail: 10 min margin; a sleep of the PC can still let the worker end first
        // (then the next search starts a new worker, or, in a rare race, reports "hata" once).
        int KeepS { get { return KeepMs <= 0 ? 0 : KeepMs / 1000 + 600; } }

        void Begin(WorkerSession s)
        {
            if (closing) return; // the window closes: no new worker
            session = s;
            UpdateButtons(); UpdateSummary();
            if (s.Mode != "status") { progress.Visible = true; progress.Style = ProgressBarStyle.Marquee; }
            var p = warm; warm = null; keepTimer.Stop();
            bool fresh = p == null || p.Exited;
            try
            {
                if (fresh) { p = new WorkerProcess(this, settings); p.Start(); }
                s.Start(p, KeepS);
            }
            catch (Exception ex)
            {
                s.Kill("start failed");
                if (p != null) p.Close();
                SetStatus("Arama bileşeni başlatılamadı: " + ex.GetType().Name);
                return;
            }
            if (fresh) { var h = WorkerStarted; if (h != null) h(s.Mode, p.StartedPid); }
        }

        /// <summary>The kept worker and its model: null (none), "loading", "ready", "error".</summary>
        public string ModelState
        {
            get { var p = session != null ? session.Proc : warm; return p != null && !p.Exited ? p.Model : null; }
        }
        public int WarmPid { get { return warm != null ? warm.Pid : 0; } } // test hook

        // Keep-alive: the worker stays for KeepMs after its last work (window open), then closes with its model.
        void ArmKeep()
        {
            keepTimer.Stop();
            if (warm == null) return;
            if (KeepMs <= 0) { CloseWarm("keep-alive 0"); return; }
            keepTimer.Interval = KeepMs;
            keepTimer.Start();
        }

        void CloseWarm(string reason)
        {
            keepTimer.Stop();
            var p = warm; warm = null;
            if (p == null) return;
            p.Close(); // stdin closes: the worker exits (the model leaves the memory)
            if (session == null && !closing) SetStatus("Yapay zekâ modeli kapatıldı (" + (reason == "keep-alive" ? "uzun süre işlem yapılmadı" : "ayar değişti") +
                                                       "). Sonraki arama modeli yeniden yükler.");
            var h = ModelStateChanged; if (h != null) h(null);
        }

        // Worker events between sessions (model preload) and of a kept worker.
        internal void OnWorkerEvent(WorkerProcess p, Dictionary<string, object> ev)
        {
            string state = ev.ContainsKey("state") ? ev["state"] as string : null;
            if (session == null && !closing)
                SetStatus(state == "loading" ? "Yapay zekâ modeli arka planda hazırlanıyor. Bu sırada arama yapabilirsiniz."
                        : state == "ready" ? "Yapay zekâ hazır: aramalar beklemeden çalışır." : "Yapay zekâ modeli yüklenemedi. Arama yine çalışır; yeniden deneyin.");
            if (p == warm && state == "ready") ArmKeep(); // the keep-alive time counts from the end of the load
            var h = ModelStateChanged; if (h != null) h(state);
        }

        public void SetKeepAliveMs(int ms) { KeepAliveMsOverride = ms; ArmKeep(); } // test hook

        internal void OnWorkerExit(WorkerProcess p)
        {
            if (p != warm) return;
            warm = null; keepTimer.Stop(); // the kept worker ended by itself (its own limit or a crash)
            var h = ModelStateChanged; if (h != null) h(null);
        }

        void StopSession()
        {
            if (session == null || session.Mode == "status") return;
            session.Cancel();
            SetStatus("Durduruluyor… Yapılan iş kaydediliyor.");
            stopBtn.Enabled = indexStopBtn.Enabled = false;
        }

        bool Busy { get { return session != null && session.Mode != "status"; } }

        void UpdateButtons()
        {
            string m = session != null ? session.Mode : null;
            searchBtn.Enabled = m != "search"; indexBtn.Enabled = !Busy; refreshBtn.Enabled = m == null;
            stopBtn.Enabled = m == "search"; indexStopBtn.Enabled = m == "index";
            storeBox.Enabled = !Busy;
            deleteBtn.Enabled = m == null; // no worker may hold the index files
            openBtn.Enabled = list.SelectedItems.Count > 0;
        }

        internal void OnEvent(string mode, Dictionary<string, object> ev)
        {
            string name = ev.ContainsKey("ev") ? ev["ev"] as string : "";
            EventNames.Add(name);
            var c = Dict(ev, "coverage");
            if (c != null) ShowCoverage(c);
            switch (name)
            {
                case "hello":
                    if (mode != "status") SetStatus("Klasörler taranıyor (silinen ve taşınan postalar)…");
                    break;
                case "progress":
                    ShowProgress(mode, ev);
                    break;
                case "results":
                    ShowResults(ev);
                    break;
                case "done":
                    if (ev.ContainsKey("peak_rss_mb") && ev["peak_rss_mb"] != null) WorkerPeakMb = Convert.ToDouble(ev["peak_rss_mb"]);
                    string stt = ev.ContainsKey("status") ? ev["status"] as string : "";
                    if (stt == "error" && mode != "status")
                        SetStatus((mode == "index" ? "İndeksleme" : "Arama") + " tamamlanamadı (" + (ev.ContainsKey("error") ? ev["error"] : "") + "). Yeniden deneyin.");
                    if (stt == "cancelled")
                        SetStatus("Durduruldu. Yapılan iş kaydedildi; sonraki arama ya da indeksleme kaldığı yerden sürer.");
                    if (stt == "ok" && mode == "index") SetStatus("İndeksleme bitti.");
                    break;
            }
        }

        internal void OnSessionEnd(WorkerSession s, string reason)
        {
            if (session != s) { if (s.KeepProcess) s.Proc.Close(); return; }
            session = null;
            // A good end keeps the worker (and its model) for the next work while the window is open.
            if (s.KeepProcess && !s.Proc.Exited && !closing && KeepMs > 0 &&
                s.Proc.Threads == settings.Threads && s.Proc.MemoryLimitMb == settings.MemoryLimitMb) { warm = s.Proc; ArmKeep(); }
            else if (s.KeepProcess) s.Proc.Close();
            progress.Visible = false; progress.Style = ProgressBarStyle.Blocks; progress.Value = 0;
            if (s.Mode == "index")
                phase.Text = reason == "ok" ? "İndeksleme bitti: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm", TR)
                           : (reason == "cancelled" ? "İndeksleme durduruldu." : "İndeksleme durdu (" + ReasonText(reason) + ").") +
                             " Kaldığı yerden sürdürmek için İndekslemeyi başlat'a basın.";
            if (s.Mode != "status" && reason != "ok" && reason != "cancelled" && !status.Text.Contains("tamamlanamadı"))
                SetStatus((s.Mode == "index" ? "İndeksleme" : "Arama") + " durdu: " + ReasonText(reason) + ". Arama işlemi kapatıldı.");
            UpdateButtons(); UpdateSummary();
            var h = SessionFinished; if (h != null) h(s.Mode, reason);
            var next = queued; queued = null;
            if (next != null && !closing) next();
        }

        static string ReasonText(string r)
        {
            return r == "hata" ? "hata" : r == "zaman aşımı" ? "zaman aşımı" : r == "iptal" ? "durduruldu" : r == "yanıt yok" ? "arama işlemi yanıt vermedi"
                 : r == "window closed" ? "pencere kapandı" : r == "cancelled" ? "durduruldu" : r;
        }

        void ShowProgress(string mode, Dictionary<string, object> ev)
        {
            string ph = ev["phase"] as string;
            int done = Convert.ToInt32(ev["done"]), total = Convert.ToInt32(ev["total"]);
            LastPhase = ph; LastDone = done; LastTotal = total;
            // Time left only for an index run: a search stops its indexing at the time limit anyway.
            string eta = mode == "index" && ev.ContainsKey("eta_s") && ev["eta_s"] != null ? " – bu aşamada kalan " + Duration(Convert.ToDouble(ev["eta_s"])) : "";
            if (total > 0) { progress.Style = ProgressBarStyle.Blocks; progress.Maximum = total; progress.Value = Math.Min(done, total); }
            string text = null;
            switch (ph)
            {
                case "scan": text = string.Format(TR, "Klasör listeleri okunuyor: {0}/{1}", done, total); break;
                case "read": text = string.Format(TR, "Yeni ve değişen postalar okunuyor: {0:N0}/{1:N0}{2}", done, total, eta); break;
                case "att": text = string.Format(TR, "Ekler okunuyor (Excel, PDF, Word, resim…): {0:N0}/{1:N0}{2}", done, total, eta); break;
                case "model":
                    modelSeen = true;
                    progress.Style = ProgressBarStyle.Marquee;
                    text = done == 0 ? "Yerel yapay zekâ modeli yükleniyor…" : "Model yüklendi.";
                    break;
                case "embed":
                    double rate = ev.ContainsKey("rate") ? Convert.ToDouble(ev["rate"]) : 0;
                    text = string.Format(TR, "Anlam indeksi: {0:N0}/{1:N0} parça{2}{3}", done, total,
                                         rate > 0 ? string.Format(TR, " ({0:0.0} parça/sn)", rate) : "", eta);
                    break;
            }
            if (text == null) return;
            if (mode == "index") phase.Text = text;
            SetStatus(text + (mode == "search" && ph != "scan" ? ". Durdurursanız kaldığı yerden sürer." : ""));
        }

        static string Duration(double sec)
        {
            if (sec < 60) return "1 dk'dan az";
            if (sec < 3600) return string.Format(TR, "~{0:0} dk", Math.Ceiling(sec / 60));
            return string.Format(TR, "~{0:0} sa {1:0} dk", Math.Floor(sec / 3600), Math.Floor(sec % 3600 / 60));
        }

        void ShowResults(Dictionary<string, object> ev)
        {
            string stage = ev["stage"] as string;
            var results = ((ArrayList)ev["results"]).Cast<Dictionary<string, object>>().ToList();
            list.BeginUpdate();
            list.Items.Clear();
            foreach (var r in results)
            {
                string date = r["date"] as string ?? "";
                DateTime dt;
                if (DateTime.TryParse(date, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) date = dt.ToString("dd.MM.yyyy HH:mm", TR);
                var it = new ListViewItem(new[] { r["subject"] as string, r["sender"] as string, date,
                                                  Conf(r["confidence"] as string), string.Join(", ", ((ArrayList)r["why"]).Cast<string>()) });
                it.Tag = r;
                if ((r["confidence"] as string) == "zayıf") it.ForeColor = SystemColors.GrayText;
                list.Items.Add(it);
            }
            list.EndUpdate();
            FitColumns(list); // a vertical scroll bar makes the list narrower
            var dw = ev.ContainsKey("date_windows") ? ((ArrayList)ev["date_windows"]).Cast<string>().ToList() : new List<string>();
            string dates = dw.Count > 0 ? " Tarih aralığı: " + string.Join("; ", dw) + "." : "";
            bool semantic = ev.ContainsKey("semantic") && (bool)ev["semantic"];
            LastSemantic = semantic;
            if (stage == "keyword")
                SetStatus((results.Count > 0 ? "Kelime eşleşmeleri gösteriliyor." : "Tam kelime eşleşmesi yok.") + " Anlamsal arama sürüyor…" + dates);
            else if (!semantic)
                SetStatus((results.Count > 0 ? results.Count + " kelime eşleşmesi." : "Kelime eşleşmesi yok.") +
                          " Anlam indeksi henüz yok: İndeks sekmesinden indekslemeyi başlatın." + dates);
            else if (results.Count == 0 || results.All(r => (r["confidence"] as string) == "zayıf"))
                SetStatus("Güçlü bir eşleşme bulunamadı. Gri satırlar yalnız zayıf benzerliktir; başka ayrıntı deneyin." + dates);
            else
                SetStatus(results.Count + " sonuç. Açmak için çift tıklayın ya da Outlook'ta Aç'a basın." + dates);
            if (list.Items.Count > 0) list.Items[0].Selected = true; else ShowPreview();
            UpdateButtons();
        }

        static string Conf(string c)
        {
            return c == "yüksek" ? "Yüksek" : c == "orta" ? "Orta" : c == "kelime" ? "Kelime" : "Zayıf";
        }

        // ------------------------------------------------------------------ coverage (İndeks tab, scope line, status bar)
        class FolderCov { public int? Total; public int Indexed, Read, Embedded, AttTotal, AttDone; public double? ScannedAt; }

        void ShowCoverage(Dictionary<string, object> c)
        {
            foreach (var f in ((ArrayList)c["folders"]).Cast<Dictionary<string, object>>())
                cov[(string)f["id"]] = new FolderCov {
                    Total = f["total"] == null ? (int?)null : Convert.ToInt32(f["total"]), Indexed = Convert.ToInt32(f["indexed"]),
                    Read = Convert.ToInt32(f["read"]), Embedded = Convert.ToInt32(f["embedded"]),
                    AttTotal = Convert.ToInt32(f["att_total"]), AttDone = Convert.ToInt32(f["att_done"]),
                    ScannedAt = f["scanned_at"] == null ? (double?)null : Convert.ToDouble(f["scanned_at"]) };
            folderList.BeginUpdate();
            foreach (ListViewItem it in folderList.Items)
            {
                FolderCov fc;
                if (!cov.TryGetValue(((FolderInfo)it.Tag).Id, out fc)) continue;
                it.SubItems[2].Text = fc.Total == null ? "–" : Pct(fc.Read, fc.Total.Value);
                it.SubItems[3].Text = fc.Total == null ? "–" : Pct(fc.Embedded, Math.Max(fc.Indexed, fc.Total.Value));
                it.SubItems[4].Text = fc.AttTotal == 0 ? "–" : string.Format(TR, "{0:N0}/{1:N0}", fc.AttDone, fc.AttTotal);
                it.SubItems[5].Text = fc.ScannedAt == null ? "henüz yok"
                    : new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(fc.ScannedAt.Value).ToLocalTime().ToString("dd.MM.yyyy HH:mm", TR);
            }
            folderList.EndUpdate();
            UpdateSummary();
        }

        static string Pct(int part, int whole)
        {
            if (whole <= 0) return "%100";
            double p = 100.0 * part / whole;
            return "%" + (p >= 99.95 || p == 0 ? p.ToString("0", TR) : p.ToString("0.0", TR));
        }

        /// <summary>Bars, scope line and status bar for the checked folders (from the last coverage data).</summary>
        void UpdateSummary()
        {
            int total = 0, read = 0, indexed = 0, emb = 0, attT = 0, attD = 0, unknown = 0, sel = 0;
            foreach (ListViewItem it in folderList.Items)
            {
                if (!it.Checked) continue;
                sel++;
                var f = (FolderInfo)it.Tag;
                FolderCov fc;
                if (!cov.TryGetValue(f.Id, out fc) || fc.Total == null) { unknown++; total += f.Count; continue; }
                total += fc.Total.Value; read += fc.Read; indexed += fc.Indexed; emb += fc.Embedded; attT += fc.AttTotal; attD += fc.AttDone;
            }
            int semWhole = Math.Max(indexed, total);
            SetBar(barRead, pctRead, read, total, string.Format(TR, "{0:N0} / {1:N0} posta", read, total));
            SetBar(barSem, pctSem, emb, semWhole, string.Format(TR, "{0:N0} / {1:N0} posta", emb, semWhole));
            SetBar(barAtt, pctAtt, attD, attT, attT == 0 ? "okunacak ek yok" : string.Format(TR, "{0:N0} / {1:N0} ek", attD, attT));
            bool complete = sel > 0 && unknown == 0 && read >= total && emb >= indexed && attD >= attT;
            string sem = Pct(emb, semWhole);
            var st = storeBox.SelectedItem as StoreItem;
            string scope = string.Format(TR, "Aranacak: {0} – {1} klasör, {2:N0} posta.", st != null ? st.Name : "?", sel, total);
            if (sel == 0) { scopeLine.Text = "Klasör seçilmedi. İndeks sekmesinden klasör seçin."; summary.Text = ""; }
            else if (session != null && session.Mode == "index")
            {
                scopeLine.Text = string.Format(TR, "İndeksleme sürüyor (anlam {0}). Ara'ya basarsanız indeksleme durur; sonra kaldığı yerden sürdürebilirsiniz.", sem);
                summary.Text = "İndeks: anlam " + sem;
            }
            else if (cov.Count == 0) { scopeLine.Text = scope; summary.Text = ""; }
            else if (complete) { scopeLine.Text = scope + " İndeks tam."; summary.Text = "İndeks: tam"; }
            else if (read == 0 && emb == 0)
            {
                scopeLine.Text = scope + " Bu klasörler henüz indekslenmedi. İndeks sekmesinde posta kutusunu ve klasörleri seçip indekslemeyi başlatın.";
                summary.Text = "İndeks: yok";
            }
            else
            {
                scopeLine.Text = scope + string.Format(TR, " İndeks eksik (okunan {0}, anlam {1}{2}): sonuçlar bu kapsamla sınırlı. Tam indeks için İndeks sekmesine gidin.",
                                                       Pct(read, total), sem, attT > 0 ? ", ek " + Pct(attD, attT) : "");
                summary.Text = "İndeks: anlam " + sem;
            }
            scopeLine.ForeColor = complete || cov.Count == 0 || Busy ? SystemColors.ControlText : (Plain ? SystemColors.ControlText : Warn);
            summary.ForeColor = complete && !Plain ? Good : SystemColors.ControlText;
        }

        void SetBar(ProgressBar bar, Label pct, int part, int whole, string detail)
        {
            bar.Value = whole <= 0 ? (part > 0 ? 1000 : 0) : (int)Math.Min(1000, Math.Round(1000.0 * part / whole));
            pct.Text = (whole <= 0 ? "–" : Pct(part, whole)) + "  (" + detail + ")";
            bar.AccessibleDescription = pct.Text;
        }

        // ------------------------------------------------------------------ preview with highlighted query words
        void ShowPreview()
        {
            preview.Clear();
            if (list.SelectedItems.Count == 0)
            {
                Append("Bir sonuç seçin: e-postanın eşleşen bölümü burada görünür, aranan sözcükler işaretlenir.", Font, SystemColors.GrayText, Color.Empty, false);
                openBtn.Enabled = false;
                return;
            }
            var r = (Dictionary<string, object>)list.SelectedItems[0].Tag;
            Add(r["subject"] as string ?? "", titleFont, SystemColors.WindowText, Marks(r, "hl_subject"));
            preview.AppendText("\n");
            string date = list.SelectedItems[0].SubItems[2].Text;
            Add((r["sender"] as string ?? "") + (date.Length > 0 ? "  •  " + date : ""), Font, SystemColors.GrayText, null);
            string to = r.ContainsKey("recipients") ? r["recipients"] as string : "";
            if (!string.IsNullOrEmpty(to)) { preview.AppendText("\n"); Add("Kime: " + to, Font, SystemColors.GrayText, null); }
            preview.AppendText("\n\n");
            string inAtt = r.ContainsKey("match_attachment") ? r["match_attachment"] as string : null;
            if (!string.IsNullOrEmpty(inAtt)) { Add("Ekte bulundu: " + inAtt, boldFont, Plain ? SystemColors.WindowText : AccentDark, null); preview.AppendText("\n"); }
            Add(r["snippet"] as string ?? "", Font, SystemColors.WindowText, Marks(r, "hl"));
            string att = r["attachments"] as string;
            if (!string.IsNullOrEmpty(att)) { preview.AppendText("\n\n"); Add("Ekler: " + att, Font, SystemColors.GrayText, null); }
            preview.SelectionStart = 0; preview.ScrollToCaret();
            openBtn.Enabled = true;
        }

        static List<int[]> Marks(Dictionary<string, object> r, string key)
        {
            if (!r.ContainsKey(key) || !(r[key] is ArrayList)) return null;
            return ((ArrayList)r[key]).Cast<ArrayList>().Select(a => new[] { Convert.ToInt32(a[0]), Convert.ToInt32(a[1]) }).ToList();
        }

        // Append text; marks = [start, length] in UTF-16 units of text (from the worker) are bold on a yellow background.
        void Add(string text, Font font, Color color, List<int[]> marks)
        {
            int pos = 0;
            foreach (var m in (marks ?? new List<int[]>()).Where(m => m[0] >= 0 && m[1] > 0).OrderBy(m => m[0]))
            {
                int a = Math.Max(pos, Math.Min(text.Length, m[0])), b = Math.Min(text.Length, m[0] + m[1]);
                if (b <= a) continue;
                Append(text.Substring(pos, a - pos), font, color, Color.Empty, false);
                Append(text.Substring(a, b - a), font, color, MarkColor, true);
                pos = b;
            }
            Append(text.Substring(pos), font, color, Color.Empty, false);
        }

        void Append(string s, Font font, Color color, Color back, bool mark)
        {
            if (s.Length == 0) return;
            int start = preview.TextLength;
            preview.AppendText(s);
            preview.Select(start, preview.TextLength - start);
            // High contrast: no colors; the match is bold + underlined.
            Font bold;
            if (mark && !markFonts.TryGetValue(font, out bold))
                markFonts[font] = bold = new Font(font, Plain ? FontStyle.Bold | FontStyle.Underline : FontStyle.Bold);
            preview.SelectionFont = mark ? markFonts[font] : font;
            preview.SelectionColor = Plain ? SystemColors.WindowText : color;
            preview.SelectionBackColor = mark && !Plain ? back : SystemColors.Window;
            preview.Select(preview.TextLength, 0);
        }

        void OpenSelected()
        {
            if (list.SelectedItems.Count == 0) return;
            var r = (Dictionary<string, object>)list.SelectedItems[0].Tag;
            try { source.Open(r["eid"] as string, lastStoreId); }
            catch (Exception ex) { SetStatus("Posta açılamadı (taşınmış ya da silinmiş olabilir): " + ex.GetType().Name); }
        }

        void SetStatus(string s) { status.Text = s; }

        static Dictionary<string, object> Dict(Dictionary<string, object> d, string key)
        {
            object v;
            return d.TryGetValue(key, out v) ? v as Dictionary<string, object> : null;
        }

        // ------------------------------------------------------------------ closing
        // The worker never outlives the window. During an index run the user confirms first (or minimizes the window).
        void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (!allowClose && e.CloseReason == CloseReason.UserClosing && session != null && session.Mode == "index" &&
                MessageBox.Show(this, "İndeksleme sürüyor. Pencere kapanırsa indeksleme durur (yapılan iş kaydedilir).\n\n" +
                                "Kapatılsın mı? İndekslemenin sürmesi için Hayır'a basın ve pencereyi simge durumuna küçültün.",
                                "HatırMail", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }
            closing = true; queued = null; // a search waiting for the end of an index run is dropped
            if (session != null) session.Stop("window closed");
            CloseWarm("window closed"); // the model never outlives the window
        }

        /// <summary>Outlook closes or unloads the add-in: close the window and stop every worker.</summary>
        public void Shutdown() { allowClose = true; Close(); }

        // Test hooks (TestHost)
        public void SetQuery(string q) { query.Text = q; }
        public int ResultCount { get { return list.Items.Count; } }
        public string ResultEid(int i) { return ((Dictionary<string, object>)list.Items[i].Tag)["eid"] as string; }
        public void SelectResult(int i) { list.Items[i].Selected = true; }
        public void ClickOpen() { OpenSelected(); }
        public void ClickStop() { StopSession(); }
        public string StatusText { get { return status.Text; } }
        public string CoverageText { get { return scopeLine.Text + " | " + pctRead.Text + " | " + pctSem.Text + " | " + pctAtt.Text; } }
        public bool IndexComplete { get { return summary.Text == "İndeks: tam"; } }
        public bool ModelLoaded { get { return modelSeen; } }
        public string SessionMode { get { return session != null ? session.Mode : null; } }
        public string LastPhase; public int LastDone, LastTotal; public bool LastSemantic; // test hooks: last progress / results
        public Panel Root { get { return root; } }
        public Size DesignClientSize { get { return new Size(S(1024), S(681)); } } // client area of the default window
        public Size MinClientSize { get { return new Size(S(804), S(521)); } }
        public void ShowTab(int i) { tabs.SelectedIndex = i; }
        public string PreviewText { get { return preview.Text; } }
        public int PreviewMarkedChars
        {
            get
            {
                int n = 0;
                for (int i = 0; i < preview.TextLength; i++) { preview.Select(i, 1); if (preview.SelectionBackColor.ToArgb() == MarkColor.ToArgb() || (Plain && preview.SelectionFont.Underline)) n++; }
                preview.Select(0, 0);
                return n;
            }
        }
        public string ResultAttachment(int i)
        {
            var r = (Dictionary<string, object>)list.Items[i].Tag;
            return r.ContainsKey("match_attachment") ? r["match_attachment"] as string : null;
        }
        public int WorkerPid { get { return session != null ? session.Pid : 0; } }
        public RichTextBox PreviewBox { get { return preview; } }
        public List<KeyValuePair<string, string>> FolderIds() { return folders.Select(f => new KeyValuePair<string, string>(f.Id, f.Path)).ToList(); }
        public void Feed(string mode, Dictionary<string, object> ev) { OnEvent(mode, ev); } // test hook: a worker event
        public string IndexFileOf(string storeId) { return IndexFile(storeId); }          // test hook
        public string QueryText { get { return query.Text; } }                            // test hooks: searches
        public SearchHistory History { get { return history; } }
        public void ReloadHistory() { history = SearchHistory.Load(settings.Data); RefreshSuggestions(); } // as at the window open
        public string[] Suggestions { get { return query.AutoCompleteCustomSource.Cast<string>().ToArray(); } }
        public List<ToolStripItem> HistoryMenuItems() { BuildHistoryMenu(); return historyMenu.Items.Cast<ToolStripItem>().ToList(); }

        class StoreItem { public string Id, Name; public override string ToString() { return Name; } }
    }

    /// <summary>One worker process (embedded Python) with its pipes, Job Object and reader thread. It runs one session,
    /// or more while the search window keeps it (keep-alive): then the model stays loaded between sessions.
    /// Events go to the current session; "model" events between sessions go to the window.</summary>
    public class WorkerProcess
    {
        readonly SearchForm form;
        readonly Settings settings;
        readonly JavaScriptSerializer js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        public readonly string Token;
        public readonly int Threads, MemoryLimitMb; // values at the start (a change needs a new process)
        Process proc;
        IntPtr job = IntPtr.Zero;
        StreamWriter writer;
        public WorkerSession Current;  // UI thread only
        public string Model;           // null (not loaded), "loading", "ready", "error" (UI thread)
        public volatile bool Exited;
        public int StartedPid;         // process id, also after the exit (tests check that it is gone)
        public int Pid { get { try { return proc != null && !proc.HasExited ? proc.Id : 0; } catch { return 0; } } }

        public WorkerProcess(SearchForm form, Settings settings)
        {
            this.form = form; this.settings = settings; Threads = settings.Threads; MemoryLimitMb = settings.MemoryLimitMb;
            var b = new byte[24];
            using (var rng = RandomNumberGenerator.Create()) rng.GetBytes(b);
            Token = BitConverter.ToString(b).Replace("-", "").ToLowerInvariant();
        }

        public void Start()
        {
            var psi = new ProcessStartInfo(settings.Python)
            {
                Arguments = "-I -B -X utf8 \"" + settings.Worker + "\" --data \"" + settings.Data + "\" --model \"" +
                            settings.Model + "\" --threads " + Threads,
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true,
                RedirectStandardOutput = true, RedirectStandardError = true,
                StandardOutputEncoding = new UTF8Encoding(false), StandardErrorEncoding = new UTF8Encoding(false),
                WorkingDirectory = Path.GetDirectoryName(settings.Worker)
            };
            psi.EnvironmentVariables["HM_SESSION"] = Token;
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["OMP_NUM_THREADS"] = Threads.ToString();
            psi.EnvironmentVariables["MKL_NUM_THREADS"] = Threads.ToString();
            psi.EnvironmentVariables["HF_HOME"] = Path.Combine(settings.Home, "data", "hf-empty");
            psi.EnvironmentVariables["ORT_DISABLE_TELEMETRY"] = "1"; // ONNX Runtime: no telemetry upload, no device id
            proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            proc.Start();
            StartedPid = proc.Id;
            Native.Confine(proc, MemoryLimitMb, ref job);
            try { proc.PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
            // Our own UTF-8 writer on the raw pipe (the worker also accepts a BOM written by Process.Start).
            writer = new StreamWriter(proc.StandardInput.BaseStream, new UTF8Encoding(false)) { AutoFlush = false, NewLine = "\n" };
            proc.ErrorDataReceived += (s, e) => { }; // drain stderr (library warnings only); never shown or stored
            proc.BeginErrorReadLine();
            new Thread(ReadLoop) { IsBackground = true, Name = "HatirMail-reader" }.Start();
        }

        public void Send(object msg)
        {
            if (Exited) return;
            try { writer.Write(js.Serialize(msg)); writer.Write("\n"); writer.Flush(); }
            catch (IOException) { } // worker is gone; the reader loop ends the session
            catch (ObjectDisposedException) { }
        }

        void ReadLoop()
        {
            try
            {
                string line;
                while ((line = proc.StandardOutput.ReadLine()) != null)
                {
                    Dictionary<string, object> ev;
                    try { ev = js.Deserialize<Dictionary<string, object>>(line); } catch { continue; }
                    form.BeginInvoke((Action)(() => Dispatch(ev)));
                }
            }
            catch { }
            // The form can be gone (Outlook closing). An exception here would end the Outlook process.
            try { form.BeginInvoke((Action)OnExit); } catch { }
        }

        // UI thread
        void Dispatch(Dictionary<string, object> ev)
        {
            string name = ev.ContainsKey("ev") ? ev["ev"] as string : "";
            if (name == "model") { Model = ev["state"] as string; form.OnWorkerEvent(this, ev); return; } // preload, between sessions
            if (name == "hello" && ev.ContainsKey("model_ready") && true.Equals(ev["model_ready"])) Model = "ready";
            if (name == "progress" && (ev["phase"] as string) == "model") Model = Convert.ToInt32(ev["done"]) == 0 ? "loading" : "ready";
            if (Current != null) Current.Handle(ev);
        }

        void OnExit()
        {
            Exited = true;
            // "done" ends the session first (same queue); here only a worker that exited without "done" (crash, memory limit).
            if (Current != null) Current.End(Current.Cancelled ? "cancelled" : "hata");
            Close();
            form.OnWorkerExit(this);
        }

        public bool WaitForExit(int ms) { try { return proc == null || proc.WaitForExit(ms); } catch { return true; } }

        public void Kill() { try { if (proc != null && !proc.HasExited) proc.Kill(); } catch { } }

        /// <summary>End the process: close stdin (a kept worker exits at once), wait up to 5 s, kill, close the job
        /// (kill-on-close: no worker process can stay alive after this).</summary>
        public void Close()
        {
            Exited = true;
            if (Model == "loading") Kill(); // a model load cannot stop early: do not block the UI thread for it
            try { writer.Close(); } catch { }
            try { if (proc != null && !proc.WaitForExit(5000)) proc.Kill(); } catch { }
            Native.CloseJob(ref job);
            try { if (proc != null) proc.Dispose(); } catch { }
            proc = null;
        }
    }

    /// <summary>One session on a worker process. Mode search: one search (indexing limited to Settings.SearchIndexBudgetS).
    /// Mode index: full indexing, started by the user, no query. Mode status: index state only (no scan, no mail
    /// reading, no model; with preload the kept worker loads the model after it).</summary>
    public class WorkerSession
    {
        const int StallS = 900; // a worker that sends nothing for 15 min is stopped (an index run has no hard limit)
        readonly SearchForm form;
        readonly IMailSource source;
        readonly Settings settings;
        readonly string storeId, storeName, query;
        readonly List<FolderInfo> scope;
        public readonly string Mode;
        public readonly bool Preload;
        public WorkerProcess Proc;
        volatile bool cancelled, ended;
        public bool Cancelled { get { return cancelled; } }
        readonly System.Windows.Forms.Timer hardTimeout = new System.Windows.Forms.Timer(), stall = new System.Windows.Forms.Timer();
        readonly Stopwatch clock = Stopwatch.StartNew();
        long lastEventMs;
        string attDir;
        int attSeq;
        Dictionary<string, string> folderOf;
        public int Pid { get { return Proc != null ? Proc.Pid : 0; } }
        public int StartedPid { get { return Proc != null ? Proc.StartedPid : 0; } }
        public bool KeepProcess; // the worker stays for the next session after a good end

        public WorkerSession(SearchForm form, IMailSource source, Settings settings, string mode, string storeId, string storeName,
                             List<FolderInfo> scope, string query, bool preload = false)
        {
            this.form = form; this.source = source; this.settings = settings; this.Mode = mode; this.Preload = preload;
            this.storeId = storeId; this.storeName = storeName; this.scope = scope; this.query = query;
        }

        /// <summary>Run on proc (a kept worker or a new one). keepAliveS > 0: the worker waits so long for the next
        /// session after this one (the window closes it earlier; this is only the worker's own safety limit).</summary>
        public void Start(WorkerProcess proc, int keepAliveS)
        {
            Proc = proc;
            proc.Current = this;
            KeepProcess = keepAliveS > 0;
            if (Mode == "search")
            {
                hardTimeout.Interval = Math.Max(60, settings.SessionTimeoutS) * 1000;
                hardTimeout.Tick += (s, e) => Kill("zaman aşımı");
                hardTimeout.Start();
            }
            stall.Interval = 30000;
            stall.Tick += (s, e) => { if (clock.ElapsedMilliseconds - Interlocked.Read(ref lastEventMs) > StallS * 1000L) Kill("yanıt yok"); };
            stall.Start();
            if (settings.Attachments && Mode != "status")
            {
                // Saved attachments live here only until the worker has read them (the worker deletes each file).
                // A folder left by a crash is deleted now and at the end of this session.
                attDir = Path.Combine(settings.Data, "tmp-att");
                try { if (Directory.Exists(attDir)) Directory.Delete(attDir, true); Directory.CreateDirectory(attDir); }
                catch { attDir = null; }
            }
            var begin = new Dictionary<string, object> {
                {"op", "begin"}, {"session", proc.Token}, {"mode", Mode}, {"query", query}, {"now", DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss")},
                {"top_k", 10}, {"idle_timeout_s", settings.IdleTimeoutS}, {"keep_alive_s", keepAliveS}, {"preload", Preload && keepAliveS > 0},
                {"attachments", Mode == "status" ? settings.Attachments : attDir != null}, {"att_dir", attDir},
                {"store", new Dictionary<string, object> { {"id", storeId}, {"name", storeName} }},
                {"folders", scope.Select(f => new Dictionary<string, object> { {"id", f.Id}, {"path", f.Path} }).ToList()} };
            if (Mode == "search") begin["index_budget_s"] = settings.SearchIndexBudgetS;
            Send(begin);
        }

        void Send(object msg)
        {
            if (ended) return;
            Interlocked.Exchange(ref lastEventMs, clock.ElapsedMilliseconds); // traffic in both directions counts
            Proc.Send(msg);
        }

        // UI thread
        public void Handle(Dictionary<string, object> ev)
        {
            if (ended) return;
            Interlocked.Exchange(ref lastEventMs, clock.ElapsedMilliseconds);
            form.OnEvent(Mode, ev);
            string name = ev["ev"] as string;
            if (name == "hello" && Mode != "status") SendManifests();
            else if (name == "need") SendItems(((ArrayList)ev["eids"]).Cast<string>().ToList());
            else if (name == "need_att") SaveAttachments((ArrayList)ev["atts"]);
            else if (name == "done")
            {
                string st = ev["status"] as string;
                End(st == "ok" ? "ok" : st == "cancelled" ? "cancelled" : "hata");
            }
        }

        // Folder item lists. Pumps UI messages between pages so that "Durdur" stays clickable.
        void SendManifests()
        {
            var map = new Dictionary<string, string>();
            foreach (var f in scope)
            {
                var page = new List<object>();
                foreach (var row in source.Manifest(f, () => cancelled))
                {
                    page.Add(row);
                    map[(string)row[0]] = f.Id;
                    if (page.Count == 500)
                    {
                        Send(new Dictionary<string, object> { {"op", "manifest"}, {"folder", f.Id}, {"items", page}, {"final", false} });
                        page = new List<object>();
                        Application.DoEvents();
                        if (cancelled || ended) return;
                    }
                }
                if (cancelled || ended) return;
                Send(new Dictionary<string, object> { {"op", "manifest"}, {"folder", f.Id}, {"items", page}, {"final", true} });
            }
            folderOf = map;
            Send(new Dictionary<string, object> { {"op", "scan_done"} });
        }

        // Read the asked mails, newest first. A search reads for at most half of its indexing time (then the
        // worker uses the rest for attachments and the semantic index); the rest waits for the next run.
        void SendItems(List<string> eids)
        {
            var batch = new List<object>();
            var failed = new List<string>();
            var sw = Stopwatch.StartNew();
            long limitMs = Mode == "search" ? settings.SearchIndexBudgetS * 500L : long.MaxValue, pumped = 0;
            bool stopped = false;
            foreach (var eid in eids)
            {
                if (cancelled || ended) return;
                if (sw.ElapsedMilliseconds >= limitMs) { stopped = true; break; }
                try { batch.Add(source.Read(eid, storeId, folderOf[eid])); }
                catch { failed.Add(eid); } // item moved/deleted during the scan, or blocked by policy
                if (batch.Count == 25)
                {
                    Send(new Dictionary<string, object> { {"op", "items"}, {"items", batch} });
                    batch = new List<object>();
                }
                if (sw.ElapsedMilliseconds - pumped >= 100) { pumped = sw.ElapsedMilliseconds; Application.DoEvents(); }
            }
            if (batch.Count > 0) Send(new Dictionary<string, object> { {"op", "items"}, {"items", batch} });
            Send(new Dictionary<string, object> { {"op", "items_done"}, {"failed", failed}, {"stopped", stopped} });
        }

        // Save the asked attachments ([eid, index, name]) into attDir. The file name is a counter, never the
        // attachment name (no path tricks, no long names). A file that cannot be saved is reported as null.
        void SaveAttachments(ArrayList atts)
        {
            var files = new List<object>();
            foreach (ArrayList a in atts)
            {
                if (cancelled || ended) return;
                string eid = (string)a[0];
                int idx = Convert.ToInt32(a[1]);
                string path = Path.Combine(attDir, (++attSeq).ToString("D6") + ".bin");
                bool ok;
                try { ok = source.SaveAttachment(eid, storeId, idx, path); }
                catch { ok = false; } // mail moved/deleted, attachment blocked by policy or IRM-protected
                if (!ok) try { File.Delete(path); } catch { }
                files.Add(new object[] { eid, idx, ok ? path : null });
                Application.DoEvents();
            }
            Application.DoEvents();
            if (cancelled || ended) return;
            Send(new Dictionary<string, object> { {"op", "att_files"}, {"files", files} });
        }

        public void Cancel()
        {
            if (ended || cancelled) return;
            cancelled = true;
            Send(new Dictionary<string, object> { {"op", "cancel"} });
            // The worker stops after the current model batch (a few seconds on a slow CPU). If not, kill it.
            var t = new System.Windows.Forms.Timer { Interval = 30000 };
            t.Tick += (s, e) => { t.Stop(); t.Dispose(); Kill("iptal"); };
            t.Start();
        }

        public void Kill(string reason)
        {
            if (ended) return;
            KeepProcess = false;
            if (Proc != null) Proc.Kill();
            End(reason);
        }

        /// <summary>Stop now (window closes): cancel, wait up to 3 s (the worker finishes the file it reads, so the
        /// file is not marked as failed), then kill.</summary>
        public void Stop(string reason)
        {
            if (ended) return;
            cancelled = true;
            KeepProcess = false;
            Send(new Dictionary<string, object> { {"op", "cancel"} });
            if (Proc != null) Proc.WaitForExit(3000);
            Kill(reason);
        }

        public void End(string reason)
        {
            if (ended) return;
            ended = true;
            hardTimeout.Stop(); stall.Stop();
            if (Proc != null)
            {
                Proc.Current = null;
                // Only a good end keeps the worker (cancel and error end the worker process itself).
                if (reason != "ok" || !KeepProcess || Proc.Exited) { KeepProcess = false; Proc.Close(); }
            }
            try { if (attDir != null && Directory.Exists(attDir)) Directory.Delete(attDir, true); } catch { } // no copy stays
            form.OnSessionEnd(this, reason);
        }
    }

    /// <summary>Windows Job Object: the worker dies with the job handle (and with Outlook); memory cap.</summary>
    static class Native
    {
        const int JobObjectExtendedLimitInformation = 9;
        const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000, JOB_OBJECT_LIMIT_PROCESS_MEMORY = 0x100,
                   JOB_OBJECT_LIMIT_ACTIVE_PROCESS = 0x8;

        [StructLayout(LayoutKind.Sequential)]
        struct BASIC { public long PerProcessUserTimeLimit, PerJobUserTimeLimit; public uint LimitFlags; public UIntPtr MinimumWorkingSetSize, MaximumWorkingSetSize; public uint ActiveProcessLimit; public UIntPtr Affinity; public uint PriorityClass, SchedulingClass; }
        [StructLayout(LayoutKind.Sequential)]
        struct IO { public ulong a, b, c, d, e, f; }
        [StructLayout(LayoutKind.Sequential)]
        struct EXT { public BASIC Basic; public IO Io; public UIntPtr ProcessMemoryLimit, JobMemoryLimit, PeakProcessMemoryUsed, PeakJobMemoryUsed; }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateJobObject(IntPtr a, string name);
        [DllImport("kernel32.dll")] static extern bool SetInformationJobObject(IntPtr job, int cls, ref EXT info, int len);
        [DllImport("kernel32.dll")] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr proc);
        [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);

        public static void Confine(Process p, int memoryMb, ref IntPtr job)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return; // TestHost on Mac/mono
            job = CreateJobObject(IntPtr.Zero, null);
            var info = new EXT();
            info.Basic.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE | JOB_OBJECT_LIMIT_PROCESS_MEMORY | JOB_OBJECT_LIMIT_ACTIVE_PROCESS;
            info.Basic.ActiveProcessLimit = 1; // the worker cannot start other processes
            info.ProcessMemoryLimit = (UIntPtr)((ulong)memoryMb * 1024 * 1024);
            SetInformationJobObject(job, JobObjectExtendedLimitInformation, ref info, Marshal.SizeOf(typeof(EXT)));
            AssignProcessToJobObject(job, p.Handle);
        }

        public static void CloseJob(ref IntPtr job)
        {
            if (job != IntPtr.Zero) { CloseHandle(job); job = IntPtr.Zero; }
        }
    }
}

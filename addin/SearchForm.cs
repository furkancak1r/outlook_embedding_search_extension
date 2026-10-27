// HatırMail window (UI thread only; Outlook objects are touched only here) and the worker session.
// Tabs: Ara (search), İndeks (index state, folders, full indexing started by the user), Ayarlar (settings.json).
// Accessibility: every control has a Turkish AccessibleName and AccessibleDescription; colors are SystemColors only, and
// high contrast follows the Windows theme while the window is open; a focus frame marks the active control; access keys
// and shortcuts (F1 lists them); one sentence per result row for a screen reader; text size 100-200 % and per-monitor DPI.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
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
        public static float TestScale; // test hook: > 0 simulates a high-DPI screen (sizes and fonts x TestScale) on a 96-DPI screen
        /// <summary>Test hook: null = the Windows setting; true or false simulates a high-contrast theme.</summary>
        public static bool? HighContrastOverride;
        /// <summary>Read at each use: the window follows a theme change while it is open (OnSystemColorsChanged).</summary>
        static bool HC { get { return HighContrastOverride ?? SystemInformation.HighContrast; } }

        readonly IMailSource source;
        readonly Settings settings;
        // k: 96-DPI pixels -> pixels of this window; pt: font points. Both = DPI factor x text size (Ayarlar). dpi: the DPI
        // of this window (WM_DPICHANGED); textPct: the text size in use.
        float k, pt;
        int dpi, textPct;
        Font titleFont, boldFont;
        bool scaling, ringQueued;
        readonly Panel root = new Panel { Dock = DockStyle.Fill };
        readonly Panel header = new Panel { Dock = DockStyle.Top }, bar = new Panel { Dock = DockStyle.Bottom }, rule = new Panel { Dock = DockStyle.Bottom };
        readonly Label title = new Label { Text = "HatırMail", AutoSize = true, Dock = DockStyle.Top }, subtitle = new Label { AutoSize = true, Dock = DockStyle.Top };
        readonly FocusFrame ring = new FocusFrame();
        readonly TabControl tabs = new TabControl();
        readonly TabPage tabSearch = new TabPage("Ara"), tabIndex = new TabPage("İndeks"), tabSettings = new TabPage("Ayarlar");
        // Ara
        readonly TableLayoutPanel qrow = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 5, RowCount = 1, AutoSize = true };
        readonly TextBox query = new TextBox();
        readonly Button searchBtn, stopBtn, openBtn, historyBtn;
        readonly ContextMenuStrip historyMenu = new ContextMenuStrip { ShowItemToolTips = true };
        SearchHistory history; // recent and saved searches (data\searches.json)
        readonly Label scopeLine = new Label();
        // Filters (Ara tab): sender and recipient values from the worker (done.facets), folder = a checked İndeks folder.
        // These three lists open with a search field (SearchList); the folder list takes more than one folder (check boxes).
        readonly ComboBox senderBox = new SearchList(), recipientBox = new SearchList();
        readonly SearchList folderBox = new SearchList(true);
        // Date range (presets or "Özel aralık…" with two date fields in a third row) and attachments; worker keys
        // date_from / date_to (ISO local time, to = exclusive) and attach.
        readonly ComboBox dateBox = new ComboBox(), attFilterBox = new ComboBox();
        readonly DateTimePicker dateFrom = new DateTimePicker(), dateTo = new DateTimePicker();
        readonly TableLayoutPanel filterTable = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 6, RowCount = 3 };
        readonly List<Label> filterLabels = new List<Label>();
        readonly Control[] filterFields;
        readonly List<Control> rangeRow = new List<Control>(); // the third filter row: only for "Özel aralık…"
        readonly Button clearFiltersBtn;
        bool fillingFilters, dateOpen;
        readonly ResultList list = new ResultList();
        readonly ContextMenuStrip resultMenu = new ContextMenuStrip();
        static readonly string[] ResultColumns = { "Konu", "Kişi", "Tarih", "Güven", "Neden" };
        int sortCol = 3; bool sortDesc = true; // 3 "Güven" descending = the worker's order (relevance); kept for new results
        Action repeatLast; // the last search or "Benzerlerini bul" that started; null = none yet (result menu)
        string searchedText; // the query box text at the start of the last search (first-result focus); null = no move
        readonly RichTextBox preview = new RichTextBox();
        readonly Panel previewBox = new Panel();                         // the attachment bar and the preview
        readonly TableLayoutPanel attBar = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, AutoSize = true, Visible = false };
        readonly Label attLine = new Label();
        readonly Button openAttBtn;
        /// <summary>Opens a saved attachment copy with its program (test hook: TestHost records the path instead).
        /// No program for the type: Windows asks for one (ErrorDialog).</summary>
        public Action<string> Launch;
        // İndeks
        readonly ComboBox storeBox = new ComboBox();
        readonly ListView folderList = new ListView();
        readonly TableLayoutPanel grid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, RowCount = 3, AutoSize = true };
        readonly ProgressBar barRead = new ProgressBar(), barSem = new ProgressBar(), barAtt = new ProgressBar();
        readonly Label pctRead = new Label(), pctSem = new Label(), pctAtt = new Label(), phase = new Label();
        readonly Button indexBtn, updateBtn, indexStopBtn, refreshBtn;
        // Ayarlar
        readonly NumericUpDown numThreads = new NumericUpDown(), numMemory = new NumericUpDown(), numBudget = new NumericUpDown(), numKeep = new NumericUpDown();
        readonly ComboBox textSizeBox = new ComboBox();
        readonly CheckBox chkAtt = new CheckBox(), chkFirst = new CheckBox();
        readonly Button saveBtn, defaultsBtn, deleteBtn, excludeBtn;
        // Status bar
        readonly Label status = new Label(), summary = new Label();
        readonly ProgressBar progress = new ProgressBar();

        List<FolderInfo> folders = new List<FolderInfo>();      // folders of the store that are indexed (İndeks list)
        List<FolderInfo> allFolders = new List<FolderInfo>();   // all mail folders of the store
        HashSet<string> excludedIds = new HashSet<string>();    // excluded in Ayarlar (with subfolders): never in a scope
        public int LastExcludedRemoved = -1; // test hook: excluded_removed of the last done
        bool purgeAsked;    // the next status check applies a new folder exclusion: it reports the deleted mails
        string notice;      // that report: it stays in front of the model texts until the next session starts
        string filterNote = ""; // the filters of the running or last search (status text)
        string autoEid;         // the result that ShowResults selected itself (the first row)
        int previewWant; bool fitQueued; // preview height: the default or the splitter; FitPreview can make it smaller
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

        /// <summary>The shortcuts that are not an access key of a control (F1 list). The access keys (Alt+letter) are read
        /// from the controls, so the list cannot miss one.</summary>
        public static readonly string[][] Shortcuts = {
            new[] { "Ctrl+F", "Arama kutusuna gider." },
            new[] { "Esc", "Arama kutusuna döner." },
            new[] { "F1", "Bu listeyi açar." },
            new[] { "Tab, Shift+Tab", "Sonraki ya da önceki alana geçer." },
            new[] { "Ctrl+Tab, Ctrl+Shift+Tab", "Sonraki ya da önceki sekmeyi açar." },
            new[] { "Alt+A, Alt+İ, Alt+Y", "Ara, İndeks ya da Ayarlar sekmesini açar." },
            new[] { "Enter (arama kutusunda)", "Aramayı başlatır." },
            new[] { "Enter (sonuç listesinde)", "Seçili e-postayı Outlook'ta açar." },
            new[] { "Shift+F10, Menü tuşu (sonuç listesinde)", "Seçili e-postanın menüsünü açar." },
            new[] { "Ctrl+1, Ctrl+2, Ctrl+3, Ctrl+4, Ctrl+5", "Sonuçları Konu, Kişi, Tarih, Güven ya da Neden sütununa göre sıralar. Yeniden basınca sıra döner." },
            new[] { "Alt+Aşağı ok (bir listede)", "Listeyi açar. Gönderen, Alıcı ve Klasör listelerinde yazarak arayabilirsiniz." } };
        static readonly char[] TabKeys = { 'A', 'İ', 'Y' }; // Alt+A Ara, Alt+İ İndeks, Alt+Y Ayarlar (a TabControl shows "&" as text)

        public SearchForm(IMailSource source, Settings settings)
        {
            this.source = source;
            this.settings = settings;
            // Sizes below are 96-DPI pixels; S() converts them to this window: DPI x text size (Ayarlar "Yazı büyüklüğü").
            // Fonts in points follow the system DPI already, so pt = window DPI / system DPI x text size. No WinForms
            // auto-scaling: in a 2x test it scaled a label in a nested panel 4x. A new DPI (a monitor with another DPI:
            // WM_DPICHANGED) or a new text size scales the open window (ApplyScale).
            float sys = TestScale > 0 ? 96 * TestScale : ScreenDpi();
            dpi = (int)Math.Round(sys);
            textPct = settings.TextScalePct;
            k = sys / 96F * textPct / 100F;
            pt = (TestScale > 0 ? TestScale : 1F) * textPct / 100F;
            Text = "HatırMail";
            Font = new Font("Segoe UI", 9f * pt);
            titleFont = new Font("Segoe UI Semibold", 11.5f * pt);
            boldFont = new Font("Segoe UI", 9f * pt, FontStyle.Bold);
            Width = S(1040); Height = S(720); MinimumSize = new Size(S(820), S(560)); AutoScroll = true; // FitToScreen
            StartPosition = FormStartPosition.CenterScreen;

            searchBtn = MakeButton("A&ra", true, "Aramayı başlatır. Kısayol: Alt+R.");
            stopBtn = MakeButton("D&urdur", false, "Çalışan aramayı durdurur. Yapılan iş kaydedilir. Kısayol: Alt+U.");
            openBtn = MakeButton("Outlook'ta A&ç", false, "Seçili e-postayı Outlook'ta açar. Kısayol: Alt+Ç.");
            historyBtn = MakeButton("Ara&malar ▼", false, "Son ve kayıtlı aramaların menüsünü açar. Kısayol: Alt+M.");
            openAttBtn = MakeButton("&Eki aç", false, "Eşleşen eki açar. Kısayol: Alt+E.");
            indexBtn = MakeButton("İndekslemeyi &başlat", true, "İşaretli klasörlerin tam indekslemesini başlatır ya da kaldığı yerden sürdürür. Kısayol: Alt+B.");
            updateBtn = MakeButton("&Güncelle", false, "Yalnız son indekslemeden sonra gelen, değişen, taşınan ve silinen postaları işler. Kısayol: Alt+G.");
            indexStopBtn = MakeButton("&Durdur", false, "Çalışan indekslemeyi ya da güncellemeyi durdurur. Yapılan iş kaydedilir. Kısayol: Alt+D.");
            refreshBtn = MakeButton("Durumu yenil&e", false, "İndeks durumunu yeniden okur. Kısayol: Alt+E.");
            saveBtn = MakeButton("&Kaydet", true, "Ayarları kaydeder. Yazı büyüklüğü hemen uygulanır. Kısayol: Alt+K.");
            defaultsBtn = MakeButton("&Varsayılan değerler", false, "Alanlara varsayılan değerleri yazar. Kaydetmek için Kaydet'e basın. Kısayol: Alt+V.");
            deleteBtn = MakeButton("İndeksi &sil", false, "Bu bilgisayardaki arama indeksini onaydan sonra siler. Kısayol: Alt+S.");
            excludeBtn = MakeButton("Klasörleri &hariç tut…", false, "İndekslenmeyecek ve aranmayacak klasörleri seçin. Kısayol: Alt+H.");
            clearFiltersBtn = MakeButton("&Filtreleri temizle", false, "Bütün filtreleri kaldırır. Kısayol: Alt+F.");
            stopBtn.AccessibleName = "Aramayı durdur"; indexStopBtn.AccessibleName = "İndekslemeyi durdur";

            // Header band
            title.Font = new Font("Segoe UI Semibold", 15f * pt);
            subtitle.Text = "Hatırladığınız birkaç ayrıntıyı yazın; e-postayı bulun ve Outlook'ta açın. Her şey bu bilgisayarda çalışır.";
            Wrap(subtitle); // large text on a small screen: two lines, not cut
            header.Controls.Add(subtitle); header.Controls.Add(title);
            subtitle.SizeChanged += (s, e) => FitHeader();

            // Status bar: the current work on the left, the index summary on the right
            status.Dock = DockStyle.Fill; status.AutoEllipsis = true; status.TextAlign = ContentAlignment.MiddleLeft;
            summary.Dock = DockStyle.Right; summary.AutoSize = true; summary.TextAlign = ContentAlignment.MiddleRight;
            progress.Dock = DockStyle.Right; progress.Visible = false;
            bar.Controls.Add(status); bar.Controls.Add(progress); bar.Controls.Add(summary);

            // ponytail: the tab padding is set once (a new value later makes a new tab window and all its child windows);
            // after a scale change the tab headers keep this padding, the text still fits (tab width = text + padding).
            tabs.Dock = DockStyle.Fill; tabs.Padding = new Point(S(14), S(5)); tabs.ShowToolTips = true; tabs.AccessibleName = "Sekmeler";
            tabs.AccessibleDescription = "Ara, İndeks ve Ayarlar sekmeleri. Ctrl+Tab sonraki sekmeyi açar.";
            tabs.TabPages.Add(tabSearch); tabs.TabPages.Add(tabIndex); tabs.TabPages.Add(tabSettings);
            var tabHelp = new[] { "Arama, filtreler ve sonuçlar.", "İndeks durumu, klasörler ve indeksleme.", "Ayarlar ve indekslenmeyecek klasörler." };
            for (int i = 0; i < 3; i++)
            {
                var p = tabs.TabPages[i];
                p.UseVisualStyleBackColor = true; p.AccessibleName = p.Text;
                p.ToolTipText = "Alt+" + TabKeys[i]; p.AccessibleDescription = tabHelp[i] + " Kısayol: Alt+" + TabKeys[i] + ".";
            }

            // ---- Ara: query row, hint, filters, scope line, results, preview
            qrow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 4; i++) qrow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            query.Font = new Font("Segoe UI", 12f * pt); query.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            query.Margin = new Padding(0, S(3), S(6), S(3)); query.AccessibleName = "Arama metni";
            query.AccessibleDescription = "Hatırladığınız ayrıntıları yazın ve Enter'a basın. Kısayol: Ctrl+F.";
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
            // Filters: one table, so that the fields line up in three columns of the same width. "(tümü)" = no filter; a
            // change applies at the next Ara. Sizes: Measure.
            //   Gönderen  [list]   Alıcı [list]   Klasör [list]
            //   Tarih     [list]   Ek    [list]           [Filtreleri temizle]
            //   Başlangıç [date]   Bitiş [date]     (this row only for "Özel aralık…")
            var cells = new[] {
                Tuple.Create("Gö&nderen:", (Control)senderBox, "Gönderen filtresi", "Yalnız bu kişinin gönderdiği e-postaları arar. Liste açılınca ad yazarak arayabilirsiniz. Kısayol: Alt+N."),
                Tuple.Create("A&lıcı:", (Control)recipientBox, "Alıcı filtresi", "Yalnız bu kişiye giden e-postaları arar. Liste açılınca ad yazarak arayabilirsiniz. Kısayol: Alt+L."),
                Tuple.Create("Kla&sör:", (Control)folderBox, "Klasör filtresi", "Yalnız işaretli klasörlerde arar. Birden çok klasör seçebilirsiniz. Kısayol: Alt+S."),
                Tuple.Create("&Tarih:", (Control)dateBox, "Tarih filtresi", "Yalnız bu tarih aralığındaki e-postaları arar. Kısayol: Alt+T."),
                Tuple.Create("E&k:", (Control)attFilterBox, "Ek filtresi", "Eki olan, olmayan ya da bir türde eki olan e-postaları arar. Kısayol: Alt+K."),
                Tuple.Create("", (Control)clearFiltersBtn, "", ""),
                Tuple.Create("&Başlangıç:", (Control)dateFrom, "Başlangıç tarihi", "Özel aralığın ilk günü. Kısayol: Alt+B."),
                Tuple.Create("Biti&ş:", (Control)dateTo, "Bitiş tarihi", "Özel aralığın son günü. Bu gün de aralığa dahildir. Kısayol: Alt+Ş.") };
            for (int i = 0; i < cells.Length; i++)
            {
                int row = i / 3, col = i % 3 * 2;
                var c = cells[i].Item2;
                c.Anchor = c is Button ? AnchorStyles.Right | AnchorStyles.Top : AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
                if (cells[i].Item1.Length > 0)
                {
                    var l = new Label { Text = cells[i].Item1, AutoSize = true, Anchor = AnchorStyles.Left };
                    filterTable.Controls.Add(l, col, row);
                    filterLabels.Add(l);
                    c.AccessibleName = cells[i].Item3; c.AccessibleDescription = cells[i].Item4;
                    if (row == 2) rangeRow.Add(l);
                }
                filterTable.Controls.Add(c, col + 1, row);
                if (row == 2) rangeRow.Add(c);
            }
            filterFields = cells.Select(x => x.Item2).ToArray();
            for (int i = 0; i < 3; i++) { filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 0)); filterTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3F)); }
            for (int i = 0; i < 3; i++) filterTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 0)); // Measure and FilterRows set the heights
            foreach (var box in AllFilterBoxes)
            {
                box.DropDownStyle = ComboBoxStyle.DropDownList; box.MaxDropDownItems = 20;
                box.Items.Add(Choice.All);
            }
            foreach (var c in new[] { "7d|Son 7 gün", "30d|Son 30 gün", "3m|Son 3 ay", "12m|Son 12 ay", "year|Bu yıl", "lastyear|Geçen yıl", "custom|Özel aralık…" })
                dateBox.Items.Add(new Choice { Key = c.Split('|')[0], Text = c.Split('|')[1] });
            foreach (var c in new[] { "yes|Ekli", "no|Eksiz", "excel|Excel", "pdf|PDF", "word|Word", "powerpoint|PowerPoint", "image|Resim", "archive|Zip, RAR" })
                attFilterBox.Items.Add(new Choice { Key = c.Split('|')[0], Text = c.Split('|')[1] });
            foreach (var p in new[] { dateFrom, dateTo })
            {
                p.Format = DateTimePickerFormat.Custom; p.CustomFormat = "dd.MM.yyyy";
                p.DropDown += (s, e) => dateOpen = true; p.CloseUp += (s, e) => dateOpen = false; // Esc closes the calendar first
            }
            dateFrom.Value = DateTime.Today.AddMonths(-1); dateTo.Value = DateTime.Today;
            clearFiltersBtn.Enabled = false;
            foreach (var c in rangeRow) c.Visible = false;
            // The start values above are not a user change: the handlers come after them.
            foreach (var box in AllFilterBoxes) { box.SelectedIndex = 0; box.SelectedIndexChanged += (s, e) => FilterChanged(); }
            folderBox.CheckedChanged += (s, e) => FilterChanged();
            foreach (var p in new[] { dateFrom, dateTo }) p.ValueChanged += (s, e) => FilterChanged();
            clearFiltersBtn.Click += (s, e) => ClearFilters();
            Wrap(scopeLine); scopeLine.Padding = new Padding(0, S(4), 0, S(8));
            scopeLine.Cursor = Cursors.Hand; scopeLine.Click += (s, e) => tabs.SelectedTab = tabIndex;
            list.Dock = DockStyle.Fill; list.View = View.Details; list.FullRowSelect = true; list.MultiSelect = false;
            list.HideSelection = false; list.AccessibleName = "Sonuçlar";
            list.AccessibleDescription = "Bulunan e-postalar. Her satır tek cümle okunur. Enter açar, Shift+F10 menüyü açar, Ctrl+1 ile Ctrl+5 sıralar.";
            // "Neden" fits its longest form ("kelime 1/2, anlam 0.61, ekte, tarih dışı"); Konu takes the rest.
            // "Güven" is wider than its texts: the sort arrow (▲/▼) is part of the header text.
            var widths = new[] { 300, 140, 115, 80, 265 };
            for (int i = 0; i < ResultColumns.Length; i++) list.Columns.Add(ResultColumns[i], S(widths[i]));
            // Owner drawn: the row text (what a screen reader reads) is one sentence; the cells show the columns (DrawCell).
            list.OwnerDraw = true;
            list.DrawColumnHeader += (s, e) => e.DrawDefault = true;
            list.DrawItem += (s, e) => { }; // Details view: DrawSubItem draws each cell
            list.DrawSubItem += DrawCell;
            // ponytail: an owner-drawn full-row list can repaint only the first cell under the mouse; the row is painted again.
            list.MouseMove += (s, e) => { var it = list.GetItemAt(e.X, e.Y); if (it != null) list.Invalidate(it.Bounds); };
            list.GotFocus += (s, e) => list.Invalidate(); list.LostFocus += (s, e) => list.Invalidate(); // selection colors, row frame
            list.Resize += (s, e) => FitColumns(list);
            list.SelectedIndexChanged += (s, e) => ShowPreview();
            list.DoubleClick += (s, e) => OpenSelected();
            // KeyDown comes with the message to the list (no pre-translation needed): it works in the Outlook message loop.
            list.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter && e.Modifiers == Keys.None) { e.Handled = true; OpenSelected(); } };
            // A click on a column header sorts by it; a second click turns the order. Güven = the worker's order.
            list.ColumnClick += (s, e) => HeaderClick(e.Column);
            list.ListViewItemSorter = new ResultOrder(this);
            ShowSortArrow();
            // Right-click on a result (or Shift+F10, the menu key: ResultList): open, open the attachment, copy the subject,
            // filter by its sender.
            resultMenu.Font = Font;
            resultMenu.Items.Add("Outlook'ta &aç", null, (s, e) => OpenSelected());
            resultMenu.Items.Add("&Eki aç", null, (s, e) => OpenAttachment());
            resultMenu.Items.Add("&Konuyu kopyala", null, (s, e) => CopySubject());
            resultMenu.Items.Add("&Yalnız bu gönderenle yeniden ara", null, (s, e) => FilterBySelectedSender());
            resultMenu.Opening += (s, e) =>
            {
                var r = SelectedResult();
                if (r == null) { e.Cancel = true; return; }
                resultMenu.Items[1].Visible = AttIndex(r) > 0;
                resultMenu.Items[3].Enabled = CanFilterBySender(r);
            };
            list.ContextMenuStrip = resultMenu;
            Disposed += (s, e) => resultMenu.Dispose();
            // Plain text only (never RTF from a mail), read-only, no links.
            preview.Dock = DockStyle.Fill; preview.ReadOnly = true; preview.DetectUrls = false;
            preview.BorderStyle = BorderStyle.FixedSingle; preview.BackColor = SystemColors.Window; preview.AccessibleName = "Önizleme";
            preview.AccessibleDescription = "Seçili e-postanın eşleşen bölümü. Aranan sözcükler kalın yazılır.";
            preview.ScrollBars = RichTextBoxScrollBars.Vertical;
            // Above the preview, only for a match inside an attachment: "Eki aç" + "Ekte bulundu: name › member".
            attBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); attBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            attBar.Padding = new Padding(0, 0, 0, S(4));
            attLine.AutoSize = false; attLine.AutoEllipsis = true; attLine.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            attLine.Font = boldFont; attLine.TextAlign = ContentAlignment.MiddleLeft;
            attBar.Controls.Add(openAttBtn, 0, 0); attBar.Controls.Add(attLine, 1, 0);
            openAttBtn.Click += (s, e) => OpenAttachment();
            // 160 (was 190): the filter row takes its place, the result list keeps its height.
            previewBox.Dock = DockStyle.Bottom; previewBox.Height = previewWant = S(160);
            previewBox.Controls.Add(preview); previewBox.Controls.Add(attBar);
            Order(attBar, preview);
            var split = new Splitter { Dock = DockStyle.Bottom, Height = S(6), TabStop = false };
            split.SplitterMoved += (s, e) => previewWant = previewBox.Height; // the user's preview height
            // The result list keeps room for its header and 3 rows: in a small window, or with the date row of the
            // filters, the preview gives way (2.19 layout test, minimum window: the list showed only its header). After
            // the layout (Resize comes from it), so that the list height is the new one.
            list.Resize += (s, e) => { if (IsHandleCreated && !fitQueued) { fitQueued = true; BeginInvoke((Action)FitPreview); } };
            tabSearch.Controls.Add(list); tabSearch.Controls.Add(split);
            tabSearch.Controls.Add(previewBox); tabSearch.Controls.Add(scopeLine); tabSearch.Controls.Add(hint); tabSearch.Controls.Add(filterTable);
            tabSearch.Controls.Add(qrow);
            Order(qrow, hint, filterTable, scopeLine, list, split, previewBox); // the order on the screen, top to bottom
            Launch = path => { using (Process.Start(new ProcessStartInfo(path) { UseShellExecute = true, ErrorDialog = true, ErrorDialogParentHandle = Handle })) { } };

            // ---- İndeks: store, bars, live phase, start/stop, folders with their state
            var srow = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, RowCount = 1, AutoSize = true };
            srow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); srow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            srow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            storeBox.DropDownStyle = ComboBoxStyle.DropDownList; storeBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            storeBox.AccessibleName = "Posta kutusu"; storeBox.AccessibleDescription = "Aranacak ve indekslenecek posta kutusu. Kısayol: Alt+P.";
            storeBox.SelectedIndexChanged += (s, e) => LoadFolders();
            srow.Controls.Add(new Label { Text = "&Posta kutusu:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, S(6), 0) }, 0, 0);
            srow.Controls.Add(storeBox, 1, 0); srow.Controls.Add(refreshBtn, 2, 0);
            grid.Padding = new Padding(0, S(8), 0, S(4));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, S(250)));
            BarRow(grid, 0, "Okunan postalar", barRead, pctRead);
            BarRow(grid, 1, "Anlam indeksi", barSem, pctSem);
            BarRow(grid, 2, "Ek dosyaları", barAtt, pctAtt);
            Wrap(phase); phase.Padding = new Padding(0, S(4), 0, S(6)); phase.Font = boldFont; phase.Text = "İndeksleme çalışmıyor.";
            var brow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = false };
            brow.Controls.Add(indexBtn); brow.Controls.Add(updateBtn); brow.Controls.Add(indexStopBtn);
            var note = Wrap(new Label { ForeColor = SystemColors.GrayText, Padding = new Padding(0, S(2), 0, S(8)),
                                        Text = "Tam indeksleme uzun sürebilir; bu sırada Outlook'u kullanabilirsiniz. Durdurursanız yapılan iş kaydedilir ve sonra kaldığı yerden sürer. Outlook kapanınca indeksleme de durur. " +
                                               "Güncelle yalnız son indekslemeden sonra gelen, değişen, taşınan ve silinen postaları işler." });
            var flabel = Wrap(new Label { Padding = new Padding(0, 0, 0, S(4)),
                                          Text = "&Klasörler (işaretli klasörler aranır ve indekslenir; Ayarlar'da hariç tutulan klasörler bu listede yoktur):" });
            folderList.Dock = DockStyle.Fill; folderList.View = View.Details; folderList.CheckBoxes = true; folderList.FullRowSelect = true;
            folderList.HideSelection = false; folderList.AccessibleName = "Klasörler";
            folderList.AccessibleDescription = "İşaretli klasörler aranır ve indekslenir. Boşluk tuşu işareti değiştirir. Kısayol: Alt+K.";
            folderList.Columns.Add("Klasör", S(330));
            foreach (var c in new[] { "Posta", "Okunan", "Anlam", "Ekler" }) folderList.Columns.Add(c, S(c == "Ekler" ? 110 : 80), HorizontalAlignment.Right);
            folderList.Columns.Add("Son tarama", S(130));
            folderList.Resize += (s, e) => FitColumns(folderList);
            // ItemChecked also comes while Windows builds the native list (first show of the İndeks tab, with visual
            // styles): some items are not there yet (Outlook, 16 Oct: NullReferenceException). The summary runs once,
            // after that message. A running search or index run keeps the folder scope of its start.
            folderList.ItemChecked += (s, e) => { if (!loadingFolders) QueueSummary(); }; // the summary also rebuilds the folder filter
            tabIndex.Controls.Add(folderList); tabIndex.Controls.Add(flabel); tabIndex.Controls.Add(note); tabIndex.Controls.Add(brow);
            tabIndex.Controls.Add(phase); tabIndex.Controls.Add(grid); tabIndex.Controls.Add(srow);
            Order(srow, grid, phase, brow, note, flabel, folderList);

            // ---- Ayarlar: saved to settings.json
            var sgrid = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 2, AutoSize = true };
            sgrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); sgrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            numThreads.Minimum = 1; numThreads.Maximum = 16;
            numMemory.Minimum = 2048; numMemory.Maximum = 32768; numMemory.Increment = 512;
            numBudget.Minimum = 0; numBudget.Maximum = 600; numBudget.Increment = 5;
            numKeep.Minimum = 0; numKeep.Maximum = 120;
            SettingRow(sgrid, 0, "İşlemci &çekirdeği:", numThreads, string.Format(TR,
                "1–16. Bu bilgisayarda {0} mantıksal çekirdek var. Çok çekirdek indekslemeyi hızlandırır ama Outlook'u yavaşlatabilir.",
                Environment.ProcessorCount));
            SettingRow(sgrid, 1, "Bellek sı&nırı (MB):", numMemory, "Arama işlemi bu sınırı aşarsa kapatılır. Yapay zekâ modeliyle indeksleme 4 çekirdekte en çok yaklaşık 2.400 MB, 16 çekirdekte yaklaşık 3.400 MB kullanır.");
            SettingRow(sgrid, 2, "Ara&mada indeksleme (sn):", numBudget,
                "Ara'ya bastığınızda yeni postaları indekslemek için en çok bu kadar beklenir; kalan iş sonraki aramaya ya da İndeks sekmesine kalır. 0: aramada indeksleme yok.");
            SettingRow(sgrid, 3, "Yapay &zekâyı hazır tut (dk):", numKeep,
                "Pencere açıkken model son işten sonra bu kadar dakika bellekte kalır; sonraki arama beklemeden çalışır. Pencere kapanınca model de kapanır. 0: her aramada yeniden yüklenir.");
            textSizeBox.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (int p in new[] { 100, 125, 150, 175, 200 }) textSizeBox.Items.Add("%" + p);
            SettingRow(sgrid, 4, "Yazı &büyüklüğü:", textSizeBox,
                "Penceredeki yazıların ve alanların büyüklüğü (%100–%200). Kaydet'e basınca hemen uygulanır.");
            chkAtt.Text = "Ek &dosyalarının içinde de ara (Excel, PDF, Word, PowerPoint, resimler ve taranmış belgeler için OCR, zip…)";
            chkAtt.AutoSize = true; chkAtt.Margin = new Padding(0, S(10), 0, S(4)); chkAtt.AccessibleName = "Ek dosyalarının içinde de ara";
            chkAtt.AccessibleDescription = "Excel, PDF, Word, PowerPoint, resim ve zip eklerinin içindeki metni de arar. Kısayol: Alt+D.";
            chkFirst.Text = "Arama bitince i&lk sonuca geç";
            chkFirst.AutoSize = true; chkFirst.Margin = new Padding(0, S(4), 0, S(10)); chkFirst.AccessibleName = "Arama bitince ilk sonuca geç";
            chkFirst.AccessibleDescription = "Arama bitince odak sonuç listesindeki ilk e-postaya geçer. Arama sırasında başka bir alana geçerseniz odak değişmez. Kısayol: Alt+L.";
            sgrid.Controls.Add(chkAtt, 0, 10); sgrid.SetColumnSpan(chkAtt, 2);
            sgrid.Controls.Add(chkFirst, 0, 11); sgrid.SetColumnSpan(chkFirst, 2);
            var sbtn = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            deleteBtn.Margin = new Padding(S(24), deleteBtn.Margin.Top, deleteBtn.Margin.Right, deleteBtn.Margin.Bottom); // apart from Kaydet
            sbtn.Controls.Add(saveBtn); sbtn.Controls.Add(defaultsBtn); sbtn.Controls.Add(excludeBtn); sbtn.Controls.Add(deleteBtn);
            sgrid.Controls.Add(sbtn, 0, 12); sgrid.SetColumnSpan(sbtn, 2);
            var info = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, S(14), 0, 0), Anchor = AnchorStyles.Left | AnchorStyles.Right,
                                   Text = "Gizlilik: Arama, ek okuma ve yapay zekâ yalnız bu bilgisayarda çalışır. E-posta ve arama metni hiçbir yere gönderilmez ve günlüğe yazılmaz.\n" +
                                          "İndeks dosyaları: " + Path.Combine(settings.Data, "index") + "\nAyarlar dosyası: " + settings.ConfigFile, UseMnemonic = false };
            sgrid.Controls.Add(info, 0, 13); sgrid.SetColumnSpan(info, 2);
            tabSettings.Controls.Add(sgrid);
            // The settings are taller than the page at the minimum window size: the page scrolls (the focus scrolls the
            // focused field into view). The other pages fill the page (result list, folder list).
            tabSettings.AutoScroll = true;

            root.Controls.Add(tabs); root.Controls.Add(rule); root.Controls.Add(bar); root.Controls.Add(header);
            root.Controls.Add(ring); ring.BringToFront(); // not docked: above the other controls, no layout effect
            Order(header, tabs, rule, bar, ring);
            Controls.Add(root);
            searchBtn.Click += (s, e) => StartSearch();
            stopBtn.Click += (s, e) => StopSession();
            indexStopBtn.Click += (s, e) => StopSession();
            openBtn.Click += (s, e) => OpenSelected();
            indexBtn.Click += (s, e) => StartIndex();
            updateBtn.Click += (s, e) => StartUpdate();
            refreshBtn.Click += (s, e) => StartStatus();
            saveBtn.Click += (s, e) => SaveSettings();
            deleteBtn.Click += (s, e) => ConfirmDeleteIndex();
            excludeBtn.Click += (s, e) => AskExcluded();
            defaultsBtn.Click += (s, e) => { ShowSettings(new Settings()); SetStatus("Varsayılan değerler yüklendi. Kaydetmek için Kaydet'e basın."); };
            FormClosing += OnClosing;
            // Keys: ProcessCmdKey (WinForms pre-translation, as Tab). Second way, for a host loop that does not pre-translate:
            // the key message itself (KeyPreview; Alt+letter: WM_SYSCOMMAND SC_KEYMENU in WndProc). A key handled by the
            // first way never comes to the second.
            KeyPreview = true;
            KeyDown += (s, e) => { if (Key(e.KeyData)) e.Handled = e.SuppressKeyPress = true; };
            // Focus frame: after each change of the active control (UpdateDefaultButton), tab, layout or scroll.
            tabs.SelectedIndexChanged += (s, e) => QueueRing();
            root.Layout += (s, e) => QueueRing();
            foreach (TabPage p in tabs.TabPages) { p.Layout += (s, e) => QueueRing(); p.Scroll += (s, e) => UpdateRing(); }
            Activated += (s, e) => UpdateRing();
            Load += (s, e) =>
            {
                FitToScreen(null, 1F); // a scaled window can be larger than a small high-DPI screen (1366x768 at 125%)
                ShowSettings(settings);
                AttachmentCopy.Clean(settings.Data); // attachment copies that a program held at the last window close
                LoadStores(); // also starts the index status check of the first store (a short worker run, no model)
                ShowPreview();
                FitColumns(list); FitColumns(folderList);
                ActiveControl = query; // the window opens with the focus in the search box
                UpdateRing();
            };
            // GDI text (as TextRenderer measures it): Outlook does not call SetCompatibleTextRenderingDefault(false). A label
            // without an access key shows "&" as a character (a store, folder or path name can have one).
            foreach (var c in AllControls(this))
            {
                var l = c as Label;
                if (l != null) { l.UseCompatibleTextRendering = false; if (Mnemonic(l.Text) == '\0') l.UseMnemonic = false; }
                if (c is ButtonBase) ((ButtonBase)c).UseCompatibleTextRendering = false;
                c.GotFocus += (s, e) => QueueRing(); // also a focus change inside a control (NumericUpDown)
            }
            ApplyColors();
            Measure();
            SetStatus("Hazır. Hatırladıklarınızı yazıp Ara'ya basın. Yapay zekâ modeli yalnız bu pencere açıkken çalışır.");
            keepTimer.Tick += (s, e) => { keepTimer.Stop(); CloseWarm("keep-alive"); };
            UpdateButtons();
        }

        public static IEnumerable<Control> AllControls(Control c)
        {
            foreach (Control x in c.Controls) { yield return x; foreach (var y in AllControls(x)) yield return y; }
        }

        /// <summary>Tab order = the order on the screen (TabIndex 0, 1, 2 … in the given order).</summary>
        static void Order(params Control[] cs) { for (int i = 0; i < cs.Length; i++) cs[i].TabIndex = i; }

        /// <summary>The text without access key marks ("&amp;Ara" -> "Ara", "&amp;&amp;" -> "&amp;").</summary>
        public static string StripAmp(string t) { return (t ?? "").Replace("&&", "\u0001").Replace("&", "").Replace("\u0001", "&"); }

        /// <summary>The access key of a control text ("Gö&amp;nderen:" -> 'n'), or '\0'.</summary>
        public static char Mnemonic(string t)
        {
            t = t ?? "";
            for (int i = 0; i < t.Length - 1; i++)
            {
                if (t[i] != '&') continue;
                if (t[i + 1] != '&') return t[i + 1];
                i++; // "&&" = one "&" character
            }
            return '\0';
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
        int Line { get { return TextRenderer.MeasureText("Ag", Font).Height; } }

        /// <summary>A button; AccessibleName = its text without "&amp;" and "▼", AccessibleDescription = desc (what it does
        /// and its access key). A primary button has the Highlight color (high contrast: a system button).</summary>
        Button MakeButton(string text, bool primary, string desc)
        {
            var b = new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, S(2), S(6), S(2)),
                                 Padding = new Padding(S(10), S(3), S(10), S(3)), UseVisualStyleBackColor = true,
                                 AccessibleName = StripAmp(text).Replace("▼", "").TrimEnd(' ', '…'), AccessibleDescription = desc };
            if (primary)
            {
                b.Font = boldFont;
                EventHandler paint = (s, e) => PrimaryColors(b);
                b.EnabledChanged += paint; b.SystemColorsChanged += paint; // a disabled flat button keeps its colors: set them
                PrimaryColors(b);
            }
            return b;
        }

        static void PrimaryColors(Button b)
        {
            // A BackColor change turns UseVisualStyleBackColor off: reset the colors first.
            if (HC) { b.ResetBackColor(); b.ResetForeColor(); b.FlatStyle = FlatStyle.Standard; b.UseVisualStyleBackColor = true; return; }
            b.FlatStyle = FlatStyle.Flat; b.UseVisualStyleBackColor = false;
            b.BackColor = b.Enabled ? SystemColors.Highlight : SystemColors.ControlLight;
            b.ForeColor = b.Enabled ? SystemColors.HighlightText : SystemColors.GrayText;
            b.FlatAppearance.BorderColor = SystemColors.HotTrack; b.FlatAppearance.MouseOverBackColor = SystemColors.HotTrack;
        }

        /// <summary>Colors that are not a control default (all SystemColors). Called at the start and when Windows changes
        /// its colors (high contrast on or off): the window follows without a restart.</summary>
        void ApplyColors()
        {
            bool hc = HC;
            header.BackColor = hc ? SystemColors.Control : SystemColors.Highlight;
            title.ForeColor = subtitle.ForeColor = hc ? SystemColors.ControlText : SystemColors.HighlightText;
            attLine.ForeColor = hc ? SystemColors.ControlText : SystemColors.HotTrack;
            rule.BackColor = SystemColors.ControlDark;
            ring.BackColor = SystemColors.Highlight;
            markFonts.Clear(); // high contrast: the marks are also underlined
            if (IsHandleCreated) { ShowPreview(); list.Invalidate(); }
        }

        protected override void OnSystemColorsChanged(EventArgs e)
        {
            base.OnSystemColorsChanged(e); // the controls first (primary buttons: their own SystemColorsChanged)
            ApplyColors();
        }

        void BarRow(TableLayoutPanel g, int row, string name, ProgressBar bar, Label pct)
        {
            var l = new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, S(4), S(10), S(4)) };
            bar.Anchor = AnchorStyles.Left | AnchorStyles.Right; bar.Height = S(18); bar.Margin = new Padding(0, S(4), S(10), S(4));
            bar.Maximum = 1000; bar.AccessibleName = name;
            pct.AutoSize = true; pct.Anchor = AnchorStyles.Left; pct.Text = "–";
            g.Controls.Add(l, 0, row); g.Controls.Add(bar, 1, row); g.Controls.Add(pct, 2, row);
        }

        void SettingRow(TableLayoutPanel g, int row, string name, Control field, string help)
        {
            field.Width = S(110); field.Margin = new Padding(0, S(6), 0, 0);
            field.AccessibleName = StripAmp(name).TrimEnd(':'); field.AccessibleDescription = help + " Kısayol: Alt+" + char.ToUpper(Mnemonic(name), TR) + ".";
            var num = field as NumericUpDown;
            if (num != null) { num.TextAlign = HorizontalAlignment.Right; num.ThousandsSeparator = false; }
            g.Controls.Add(new Label { Text = name, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, S(8), S(12), 0) }, 0, row * 2);
            g.Controls.Add(field, 1, row * 2);
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

        // ------------------------------------------------------------------ scale: text size, DPI, small screens
        /// <summary>Sizes that come from the text: called at the end of the constructor and after each scale change
        /// (ApplyScale). The other sizes are S() values that Control.Scale changes with the window.</summary>
        void Measure()
        {
            int line = Line;
            header.Padding = new Padding(S(16), S(8), S(16), S(10));
            FitHeader();
            bar.Padding = new Padding(S(12), S(4), S(12), S(4)); bar.Height = line + S(12);
            summary.Padding = new Padding(S(12), 0, 0, 0);
            rule.Height = Math.Max(1, S(1));
            progress.Width = S(160);
            foreach (TabPage p in tabs.TabPages) p.Padding = new Padding(S(12));
            attLine.Height = line;
            grid.ColumnStyles[2].Width = S(250);
            dateFrom.CalendarFont = dateTo.CalendarFont = Font;
            // Filters: the label columns have a fixed width (the fields do not move when the third row shows); fixed row
            // heights from the real control heights; each field starts at a fixed distance from the row top. The table
            // lays out a list with the height of a new list, centered in the row; Windows then makes the list taller, and
            // the row cut it at the bottom (2.19, 2x test). The table height is set, not AutoSize: FilterRows adds row 3.
            var labelW = new int[3];
            foreach (var l in filterLabels)
            {
                int col = filterTable.GetColumn(l);
                l.Margin = new Padding(col == 0 ? 0 : S(14), 0, S(6), 0);
                labelW[col / 2] = Math.Max(labelW[col / 2], l.PreferredWidth + l.Margin.Horizontal);
            }
            for (int i = 0; i < 3; i++) filterTable.ColumnStyles[i * 2].Width = labelW[i];
            clearFiltersBtn.Padding = new Padding(S(10), 0, S(10), 0); // as high as a list
            filterTable.Padding = new Padding(0, S(4), 0, 0);
            int rowH = new[] { senderBox.PreferredHeight, dateFrom.PreferredHeight, clearFiltersBtn.PreferredSize.Height }.Max() + S(6);
            foreach (var c in filterFields)
                c.Margin = new Padding(0, (rowH - (c is ComboBox ? ((ComboBox)c).PreferredHeight : c is DateTimePicker ? ((DateTimePicker)c).PreferredHeight : c.PreferredSize.Height)) / 2, 0, 0);
            filterTable.RowStyles[0].Height = rowH; filterTable.RowStyles[1].Height = rowH;
            FilterRows();
        }

        void FitHeader()
        {
            int h = header.Padding.Vertical + title.PreferredHeight + WrapHeight(subtitle) + S(2);
            if (header.Height != h) header.Height = h;
        }

        /// <summary>Scale the open window by r: a new text size (Kaydet) or a new DPI (WM_DPICHANGED: newDpi, want = the
        /// window rectangle that Windows suggests). Fonts, sizes, margins, column widths and the measured sizes change;
        /// the focus stays.</summary>
        void ApplyScale(float r, int newDpi, Rectangle? want)
        {
            if (scaling) return;
            dpi = newDpi;
            if (Math.Abs(r - 1) < 0.001f) return;
            scaling = true;
            var focus = ActiveControl;
            try
            {
                foreach (var sl in new[] { (SearchList)senderBox, (SearchList)recipientBox, folderBox }) sl.ClosePopup();
                historyMenu.Close(); resultMenu.Close();
                foreach (var b in AllFilterBoxes.Concat(new[] { storeBox, textSizeBox })) if (b.DroppedDown) b.DroppedDown = false;
                k *= r; pt *= r;
                SuspendLayout();
                // The controls with their own font (bold, title, search box); the others take the window font.
                var own = AllControls(this).Where(c => !ReferenceEquals(c.Font, c.Parent.Font)).ToList();
                Font oldBold = boldFont;
                Font = Scaled(Font, r); titleFont = Scaled(titleFont, r); boldFont = Scaled(boldFont, r);
                foreach (var c in own) c.Font = ReferenceEquals(c.Font, oldBold) ? boldFont : Scaled(c.Font, r);
                historyMenu.Font = resultMenu.Font = Font;
                markFonts.Clear();
                root.Scale(new SizeF(r, r)); // bounds, margins and paddings of all controls
                foreach (var lv in new[] { (ListView)list, folderList }) foreach (ColumnHeader h in lv.Columns) h.Width = (int)Math.Round(h.Width * r);
                foreach (var b in AllFilterBoxes) b.DropDownWidth = Math.Max(1, (int)Math.Round(b.DropDownWidth * r));
                previewWant = (int)Math.Round(previewWant * r);
                Measure();
                ResumeLayout(true);
                FitToScreen(want, r);
            }
            finally { scaling = false; }
            FitColumns(list); FitColumns(folderList);
            if (IsHandleCreated) { FitPreview(); ShowPreview(); }
            if (focus != null && !focus.IsDisposed && focus.CanSelect) ActiveControl = focus;
            UpdateRing();
        }

        static Font Scaled(Font f, float r) { return new Font(f.FontFamily, f.Size * r, f.Style, f.Unit); }

        /// <summary>A new text size (Ayarlar, Kaydet).</summary>
        void ApplyText(int pct)
        {
            float r = pct / (float)textPct;
            textPct = pct;
            ApplyScale(r, dpi, null);
        }

        /// <summary>Minimum and window size inside the work area; want = the rectangle of WM_DPICHANGED; r = the scale change.
        /// A screen smaller than the minimum window (a small high-DPI screen, or a large text size): the window scrolls and
        /// keeps its layout (AutoScrollMinSize = the client area of the minimum window); nothing is cut.</summary>
        void FitToScreen(Rectangle? want, float r)
        {
            var wa = Screen.FromControl(this).WorkingArea;
            MinimumSize = new Size(Math.Min(S(820), wa.Width), Math.Min(S(560), wa.Height));
            AutoScrollMinSize = new Size(Math.Max(0, S(820) - (Width - ClientSize.Width)), Math.Max(0, S(560) - (Height - ClientSize.Height)));
            if (want != null) Bounds = want.Value;
            else if (WindowState == FormWindowState.Normal)
                Size = new Size(Math.Min((int)Math.Round(Width * r), wa.Width), Math.Min((int)Math.Round(Height * r), wa.Height));
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            // A per-monitor window (Connect: DpiScope) that opens on a monitor with another DPI than the system DPI.
            int now = TestScale > 0 ? dpi : Native.WindowDpi(Handle);
            if (now > 0 && now != dpi) ApplyScale((float)now / dpi, now, null);
        }

        [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x02E0) // WM_DPICHANGED: the window moved to a monitor with another DPI (per-monitor window only)
            {
                int now = (int)(m.WParam.ToInt64() & 0xFFFF);
                var rc = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                if (now > 0 && now != dpi) ApplyScale((float)now / dpi, now, Rectangle.FromLTRB(rc.Left, rc.Top, rc.Right, rc.Bottom));
                m.Result = IntPtr.Zero;
                return;
            }
            // WM_SYSCOMMAND SC_KEYMENU with a character: an Alt+letter that the pre-translation did not take (DefWindowProc
            // sends it to the window when WM_SYSCHAR reaches a control): the access keys of the window.
            if (m.Msg == 0x112 && ((long)m.WParam & 0xFFF0) == 0xF100 && m.LParam != IntPtr.Zero && ProcessMnemonic((char)(long)m.LParam)) return;
            base.WndProc(ref m);
        }

        // ------------------------------------------------------------------ focus frame and keyboard
        /// <summary>The focus frame: a Highlight border around the active control (a tab: its header), above the other
        /// controls. Mouse clicks go through it. Public for the TestHost screenshots (DrawToBitmap ignores its region).</summary>
        public sealed class FocusFrame : Control
        {
            public int Thickness;
            public FocusFrame() { SetStyle(ControlStyles.Selectable, false); SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true); TabStop = false; Visible = false; }
            // Only the border is painted (on screen the region also limits it; a screenshot ignores the region).
            protected override void OnPaintBackground(PaintEventArgs e) { }
            protected override void OnPaint(PaintEventArgs e)
            {
                int w = Thickness;
                using (var b = new SolidBrush(BackColor))
                {
                    e.Graphics.FillRectangle(b, 0, 0, Width, w); e.Graphics.FillRectangle(b, 0, Height - w, Width, w);
                    e.Graphics.FillRectangle(b, 0, 0, w, Height); e.Graphics.FillRectangle(b, Width - w, 0, w, Height);
                }
            }
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == 0x84) { m.Result = (IntPtr)(-1); return; } // WM_NCHITTEST: HTTRANSPARENT
                base.WndProc(ref m);
            }
        }

        protected override void UpdateDefaultButton() { base.UpdateDefaultButton(); QueueRing(); } // the active control changed

        void QueueRing()
        {
            if (ringQueued || !IsHandleCreated || IsDisposed) return;
            ringQueued = true;
            BeginInvoke((Action)(() => { ringQueued = false; UpdateRing(); })); // after the layout: the final positions
        }

        void UpdateRing()
        {
            if (IsDisposed || scaling) return;
            var a = ActiveControl;
            var t = Rectangle.Empty;
            if (a != null && a != ring && a.Visible && root.Contains(a))
                t = root.RectangleToClient(a == tabs ? tabs.RectangleToScreen(tabs.GetTabRect(Math.Max(0, tabs.SelectedIndex)))
                                                     : a.Parent.RectangleToScreen(a.Bounds));
            int w = Math.Max(2, S(3));
            var r = Rectangle.Inflate(t, w, w);
            r.Intersect(root.ClientRectangle);
            if (a != null && a != tabs && tabs.Contains(a)) r.Intersect(root.RectangleToClient(tabs.RectangleToScreen(tabs.DisplayRectangle)));
            if (t.IsEmpty || r.Width <= 2 * w || r.Height <= 2 * w) { ring.Visible = false; return; }
            if (ring.Bounds != r || ring.Thickness != w)
            {
                ring.Bounds = r; ring.Thickness = w; ring.Invalidate();
                var reg = new Region(new Rectangle(Point.Empty, r.Size));
                reg.Exclude(new Rectangle(w, w, r.Width - 2 * w, r.Height - 2 * w));
                var old = ring.Region; ring.Region = reg; if (old != null) old.Dispose();
            }
            if (root.Controls.GetChildIndex(ring) != 0) ring.BringToFront();
            ring.Visible = true;
        }

        /// <summary>Ara tab, focus in the search box, its text selected (Ctrl+F, Esc, ribbon button).</summary>
        public void FocusQuery()
        {
            tabs.SelectedTab = tabSearch;
            ActiveControl = query;
            query.SelectAll();
            UpdateRing();
        }

        bool MenuOpen
        {
            get { return historyMenu.Visible || resultMenu.Visible || new[] { (SearchList)senderBox, (SearchList)recipientBox, folderBox }.Any(x => x.PopupOpen); }
        }

        /// <summary>The window shortcuts (ProcessCmdKey and KeyDown with KeyPreview; also a test hook: PressKey). Enter in
        /// the result list: list.KeyDown; Shift+F10 and the menu key: ResultList. In Outlook the keys come through the same
        /// WinForms pre-translation as Tab (the user tested Tab in Outlook, 10-26 15:1x).</summary>
        bool Key(Keys k)
        {
            if (MenuOpen) return false;
            Action a = null;
            Keys code = k & Keys.KeyCode, mods = k & Keys.Modifiers;
            if (k == (Keys.Control | Keys.F)) a = FocusQuery;
            else if (k == Keys.F1) a = ShowShortcuts;
            else if (k == Keys.Escape && ActiveControl != query && !dateOpen && !(ActiveControl is ComboBox && ((ComboBox)ActiveControl).DroppedDown)) a = FocusQuery;
            else if (mods == Keys.Control && tabs.SelectedTab == tabSearch)
            {
                int n = code >= Keys.D1 && code <= Keys.D5 ? code - Keys.D1 : code >= Keys.NumPad1 && code <= Keys.NumPad5 ? code - Keys.NumPad1 : -1;
                if (n >= 0) a = () => HeaderClick(n);
            }
            if (a == null) return false;
            a();
            return true;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData) { return Key(keyData) || base.ProcessCmdKey(ref msg, keyData); }

        static int TabFor(char ch) { return Array.IndexOf(TabKeys, char.ToUpper(ch, TR)); }

        /// <summary>Alt+A, Alt+İ, Alt+Y open a tab (the focus goes to its header); the other access keys: the labels and
        /// buttons of the visible page.</summary>
        protected override bool ProcessMnemonic(char ch)
        {
            int t = TabFor(ch);
            if (t < 0) return base.ProcessMnemonic(ch);
            tabs.SelectedIndex = t;
            ActiveControl = tabs;
            UpdateRing();
            return true;
        }

        /// <summary>Test hook: shows a modal dialog (F1) instead of ShowDialog (a modal dialog would stop the check).</summary>
        public Action<Form> ShowModal;

        void ShowShortcuts()
        {
            using (var d = new ShortcutDialog(this)) { if (ShowModal != null) ShowModal(d); else d.ShowDialog(this); }
        }

        /// <summary>The F1 list: the shortcuts, then the access keys of each tab (read from the controls).</summary>
        public string ShortcutText()
        {
            var sb = new StringBuilder();
            sb.AppendLine("Her yerde:");
            foreach (var s in Shortcuts) sb.AppendLine("  " + s[0] + ": " + s[1]);
            foreach (TabPage p in tabs.TabPages)
            {
                sb.AppendLine();
                sb.AppendLine(p.Text + " sekmesi (Alt tuşuyla birlikte):");
                foreach (var c in AllControls(p).Where(c => c is Label && ((Label)c).UseMnemonic || c is ButtonBase))
                {
                    char m = Mnemonic(c.Text);
                    if (m != '\0') sb.AppendLine("  Alt+" + char.ToUpper(m, TR) + ": " + KeyName(c));
                }
            }
            return sb.ToString().TrimEnd();
        }

        static string KeyName(Control c)
        {
            string t = c is Label ? StripAmp(c.Text) : c.AccessibleName ?? StripAmp(c.Text);
            int cut = t.IndexOf(" (", StringComparison.Ordinal);
            if (cut > 0) t = t.Substring(0, cut);
            return t.TrimEnd(':', ' ', '…');
        }

        /// <summary>F1: the keyboard shortcuts of the window in a read-only text box (a screen reader reads it line by line).</summary>
        public sealed class ShortcutDialog : Form
        {
            public readonly TextBox Box = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, WordWrap = true };
            public readonly TableLayoutPanel Body = new TableLayoutPanel { AutoSize = true, ColumnCount = 1 };

            public ShortcutDialog(SearchForm o)
            {
                Text = "Klavye kısayolları"; Font = o.Font; AutoScaleMode = AutoScaleMode.None; FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = MinimizeBox = false; ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent;
                Box.Text = o.ShortcutText().Replace("\r\n", "\n").Replace("\n", "\r\n");
                Box.AccessibleName = "Klavye kısayolları";
                Box.AccessibleDescription = "Bu penceredeki kısayollar. Satırları ok tuşlarıyla okuyun. Esc listeyi kapatır.";
                Box.BackColor = SystemColors.Window; // a read-only box is gray by default
                Box.Width = o.S(600); Box.Height = o.S(420); Box.Margin = new Padding(0, 0, 0, o.S(12));
                var close = o.MakeButton("&Kapat", true, "Listeyi kapatır.");
                close.DialogResult = DialogResult.OK; close.Anchor = AnchorStyles.Right; close.UseCompatibleTextRendering = false;
                Body.Controls.Add(Box); Body.Controls.Add(close);
                Body.Location = new Point(o.S(14), o.S(12));
                Controls.Add(Body);
                var ps = Body.GetPreferredSize(Size.Empty);
                ClientSize = new Size(ps.Width + o.S(28), ps.Height + o.S(22));
                AcceptButton = close; CancelButton = close;
                ActiveControl = Box;
            }
        }

        // ------------------------------------------------------------------ result rows
        /// <summary>The result list. A keyboard menu request (Shift+F10, the menu key: WM_CONTEXTMENU with lParam -1) opens
        /// the menu under the selected row, not in the middle of the list.</summary>
        sealed class ResultList : ListView
        {
            public ResultList() { DoubleBuffered = true; }
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == 0x7B && unchecked((int)(long)m.LParam) == -1 && ContextMenuStrip != null) // WM_CONTEXTMENU from the keyboard (x, y = -1)
                {
                    if (SelectedItems.Count > 0)
                    {
                        var it = SelectedItems[0];
                        it.EnsureVisible();
                        var b = it.Bounds;
                        ContextMenuStrip.Show(this, new Point(Math.Max(0, b.Left) + b.Height, b.Bottom));
                    }
                    return;
                }
                base.WndProc(ref m);
            }
        }

        static string Cell(ListViewItem it, int col) { return ((string[])((Dictionary<string, object>)it.Tag)["cells"])[col]; }

        /// <summary>A result cell: the text of the column (Tag "cells"), SystemColors; the focused row has a frame.</summary>
        void DrawCell(object s, DrawListViewSubItemEventArgs e)
        {
            var it = e.Item;
            if (it == null || !(it.Tag is Dictionary<string, object>)) { e.DrawDefault = true; return; }
            bool sel = it.Selected, focus = list.Focused;
            Color back = sel ? (focus ? SystemColors.Highlight : SystemColors.Control) : SystemColors.Window;
            Color fore = sel ? (focus ? SystemColors.HighlightText : SystemColors.ControlText) : it.ForeColor;
            var cell = e.Bounds;
            if (e.ColumnIndex == 0) cell.Width = list.Columns[0].Width; // the bounds of column 0 = the whole row
            using (var b = new SolidBrush(back)) e.Graphics.FillRectangle(b, cell);
            TextRenderer.DrawText(e.Graphics, Cell(it, e.ColumnIndex), list.Font, Rectangle.Inflate(cell, -S(4), 0), fore,
                                  TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            if (it.Focused && focus)
            {
                var row = it.Bounds;
                var g = e.Graphics;
                var clip = g.Clip;
                g.SetClip(cell);
                using (var p = new Pen(sel ? SystemColors.HighlightText : SystemColors.Highlight, Math.Max(2, S(2))) { Alignment = PenAlignment.Inset })
                    g.DrawRectangle(p, row.X, row.Y, row.Width - 1, row.Height - 1);
                g.Clip = clip;
            }
        }

        string FolderPath(string id)
        {
            var f = id == null ? null : allFolders.FirstOrDefault(x => x.Id == id);
            return f != null ? f.Path : null;
        }

        // ------------------------------------------------------------------ settings
        void ShowSettings(Settings s)
        {
            numThreads.Value = Math.Max(numThreads.Minimum, Math.Min(numThreads.Maximum, s.Threads));
            numMemory.Value = Math.Max(numMemory.Minimum, Math.Min(numMemory.Maximum, s.MemoryLimitMb));
            numBudget.Value = Math.Max(numBudget.Minimum, Math.Min(numBudget.Maximum, s.SearchIndexBudgetS));
            numKeep.Value = Math.Max(numKeep.Minimum, Math.Min(numKeep.Maximum, s.KeepAliveMin));
            chkAtt.Checked = s.Attachments;
            chkFirst.Checked = s.FocusFirstResult;
            textSizeBox.SelectedIndex = Math.Max(0, Math.Min(4, (Settings.ClampTextScale(s.TextScalePct) - 100) / 25));
        }

        void SaveSettings()
        {
            var old = new[] { settings.Threads, settings.MemoryLimitMb, settings.SearchIndexBudgetS, settings.KeepAliveMin, settings.TextScalePct };
            bool oldAtt = settings.Attachments, oldFirst = settings.FocusFirstResult;
            settings.Threads = (int)numThreads.Value; settings.MemoryLimitMb = (int)numMemory.Value;
            settings.SearchIndexBudgetS = (int)numBudget.Value; settings.Attachments = chkAtt.Checked; settings.KeepAliveMin = (int)numKeep.Value;
            settings.FocusFirstResult = chkFirst.Checked; settings.TextScalePct = 100 + 25 * Math.Max(0, textSizeBox.SelectedIndex);
            try
            {
                settings.Save();
                SetStatus(session != null ? "Ayarlar kaydedildi. Yeni değerler bir sonraki işte geçerli olur." : "Ayarlar kaydedildi.");
                // A kept worker has the old threads and memory limit (set at its start): it closes; keep-alive 0 closes it too.
                if (old[0] != settings.Threads || old[1] != settings.MemoryLimitMb || settings.KeepAliveMin == 0) CloseWarm("settings");
                else if (old[3] != settings.KeepAliveMin && warm != null) ArmKeep();
                if (oldAtt != settings.Attachments) StartStatus(); // attachment counts change
                if (settings.TextScalePct != textPct) ApplyText(settings.TextScalePct); // at once, the focus stays on Kaydet
            }
            catch (Exception ex)
            {
                settings.Threads = old[0]; settings.MemoryLimitMb = old[1]; settings.SearchIndexBudgetS = old[2]; settings.KeepAliveMin = old[3];
                settings.TextScalePct = old[4]; settings.Attachments = oldAtt; settings.FocusFirstResult = oldFirst;
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
                Box.Width = o.S(380); Box.MaxLength = 80; Box.Text = name; Box.AccessibleName = "Aramanın adı"; Box.AccessibleDescription = "Aramalar menüsünde görünecek ad. En çok 80 karakter."; Box.Margin = new Padding(0, 0, 0, o.S(12));
                var ok = o.MakeButton("&Kaydet", true, "Aramayı bu adla kaydeder."); var cancel = o.MakeButton("İptal", false, "Kaydetmeden kapatır. Kısayol: Esc.");
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
            folders = new List<FolderInfo>(); allFolders = new List<FolderInfo>(); excludedIds = new HashSet<string>();
            UpdateButtons();
            if (st == null) return;
            try
            {
                allFolders = source.Folders(st.Id);
                // Folders excluded in Ayarlar (or by the default names) are not in the list: never searched, never indexed.
                excludedIds = settings.ExcludedIn(st.Id, allFolders);
                folders = allFolders.Where(f => !excludedIds.Contains(f.Id)).ToList();
                loadingFolders = true;
                folderList.BeginUpdate();
                foreach (var f in folders)
                    folderList.Items.Add(new ListViewItem(new[] { f.Path.TrimStart('\\'), f.Count.ToString("N0", TR), "", "", "", "" }) { Checked = true, Tag = f });
                folderList.EndUpdate();
                FitColumns(folderList);
                loadingFolders = false;
                UpdateSummary(); UpdateButtons();
                StartStatus(); // also deletes the excluded folders from the index (begin.exclude)
            }
            catch (Exception ex) { folderList.EndUpdate(); loadingFolders = false; SetStatus("Klasörler okunamadı: " + ex.GetType().Name); }
        }

        // ------------------------------------------------------------------ folders excluded from the index (Ayarlar)
        void AskExcluded()
        {
            var st = storeBox.SelectedItem as StoreItem;
            if (session != null || st == null || allFolders.Count == 0) return;
            using (var d = new ExcludeDialog(this, st.Name, allFolders, excludedIds))
                if (d.ShowDialog(this) == DialogResult.OK) SetExcluded(d.Chosen());
        }

        /// <summary>Save the excluded folders of the selected store (folder entry ids; subfolders follow) and apply them:
        /// the folder list is built again, and its status check deletes the mails of these folders from the index.
        /// A write error keeps the old choice. Returns false when nothing changed.</summary>
        public bool SetExcluded(List<string> ids)
        {
            var st = storeBox.SelectedItem as StoreItem;
            if (st == null) return false;
            if (session != null) { SetStatus("Önce çalışan aramayı ya da indekslemeyi durdurun."); return false; }
            List<string> old;
            bool had = settings.ExcludedFolders.TryGetValue(st.Id, out old);
            settings.ExcludedFolders[st.Id] = ids.Distinct().ToList();
            try { settings.Save(); }
            catch (Exception ex)
            {
                if (had) settings.ExcludedFolders[st.Id] = old; else settings.ExcludedFolders.Remove(st.Id);
                SetStatus("Klasör seçimi kaydedilemedi (" + ex.GetType().Name + "). Eski seçim geçerli.");
                return false;
            }
            // Results of the old scope can show mails that are not in the index any more.
            list.Items.Clear(); ShowPreview();
            LastExcludedRemoved = -1;
            purgeAsked = File.Exists(settings.Python) && File.Exists(settings.Worker); // else no status check runs
            LoadFolders();
            SetStatus(purgeAsked ? "Klasör seçimi kaydedildi. Hariç tutulan klasörlerin postaları indeksten siliniyor…"
                                 : "Klasör seçimi kaydedildi. Postalar, arama bileşeni bulununca indeksten silinir.");
            return true;
        }

        /// <summary>"İndekslenmeyecek klasörler": all mail folders of the store with check boxes. A checked folder checks
        /// its subfolders (they are excluded with it), and they cannot be unchecked alone. Kaydet returns the top checked
        /// folders only (Chosen); subfolders made later are then excluded too.</summary>
        public sealed class ExcludeDialog : Form
        {
            public readonly CheckedListBox Box = new CheckedListBox();
            public readonly TableLayoutPanel Body = new TableLayoutPanel { AutoSize = true, ColumnCount = 1 };
            readonly List<FolderInfo> all;
            bool setting;

            public ExcludeDialog(SearchForm o, string storeName, List<FolderInfo> all, HashSet<string> excluded)
            {
                this.all = all;
                Text = "İndekslenmeyecek klasörler"; Font = o.Font; AutoScaleMode = AutoScaleMode.None; FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = MinimizeBox = false; ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent;
                var l = new Label { AutoSize = true, MaximumSize = new Size(o.S(520), 0), Margin = new Padding(0, 0, 0, o.S(8)), UseCompatibleTextRendering = false, UseMnemonic = false,
                                    Text = "Posta kutusu: " + storeName + "\n\nİşaretli klasörler ve alt klasörleri indekslenmez ve aranmaz. Bu klasörlerin postaları " +
                                           "kaydettiğinizde indeksten silinir. Outlook'taki postalarınız değişmez." };
                Box.CheckOnClick = true; Box.IntegralHeight = false; Box.Width = o.S(520); Box.Height = o.S(300); Box.AccessibleName = "Hariç tutulacak klasörler"; Box.AccessibleDescription = "İşaretli klasörler ve alt klasörleri indekslenmez ve aranmaz. Boşluk tuşu işareti değiştirir.";
                Box.Margin = new Padding(0, 0, 0, o.S(12));
                foreach (var f in all) Box.Items.Add(f.Path.TrimStart('\\'), excluded.Contains(f.Id));
                Box.ItemCheck += OnItemCheck;
                var ok = o.MakeButton("&Kaydet", true, "Seçimi kaydeder. İşaretli klasörlerin postaları indeksten silinir."); var cancel = o.MakeButton("İptal", false, "Kaydetmeden kapatır. Kısayol: Esc.");
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
            }

            bool Under(int i, int parent) { return all[i].Path.StartsWith(all[parent].Path + "\\", StringComparison.Ordinal); }

            // Checked: the subfolders follow. Unchecked: refused while a parent folder is checked (the parent excludes it).
            void OnItemCheck(object s, ItemCheckEventArgs e)
            {
                if (setting) return;
                if (e.NewValue == CheckState.Unchecked && Enumerable.Range(0, all.Count).Any(p => p != e.Index && Box.GetItemChecked(p) && Under(e.Index, p)))
                { e.NewValue = CheckState.Checked; return; }
                setting = true;
                for (int i = 0; i < all.Count; i++) if (Under(i, e.Index)) Box.SetItemChecked(i, e.NewValue == CheckState.Checked);
                setting = false;
            }

            /// <summary>The checked folders without a checked parent (entry ids).</summary>
            public List<string> Chosen()
            {
                var on = Enumerable.Range(0, all.Count).Where(Box.GetItemChecked).ToList();
                return on.Where(i => !on.Any(p => p != i && Under(i, p))).Select(i => all[i].Id).ToList();
            }
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

        // ------------------------------------------------------------------ filters (Ara tab)
        /// <summary>An item of a filter list: Key = the value for the worker (null = "(tümü)").</summary>
        sealed class Choice
        {
            public static readonly Choice All = new Choice { Text = "(tümü)" };
            public string Key, Text;
            public override string ToString() { return Text; }
        }

        static string KeyOf(ComboBox b) { var c = b.SelectedItem as Choice; return c != null ? c.Key : null; }

        /// <summary>A filter list that opens with a search field (Gönderen, Alıcı, Klasör: many names and e-mail addresses).
        /// Each word in the field must be in the item text, without case and Turkish marks ("cakir" finds "Çakır").
        /// Up/Down/PageUp/PageDown move, Enter or a click takes the item, Esc closes. A letter on the closed list opens it
        /// with that letter. Item 0 is "(tümü)": it shows only while the field is empty.
        /// Multi (Klasör): check boxes; a click or Enter turns a row on or off and the list stays open; the keys are in
        /// Checked. "(tümü)" is on when no row is on; turning it on clears the rows. Item 0 of the closed box shows the
        /// choice: "(tümü)", the name, or "3 klasör seçili" (ponytail: the word "klasör": only the folder list is Multi).
        /// ponytail: a WinForms ToolStripDropDown (it closes on an outside click, also in the Outlook message loop); the
        /// native list stays for a screen reader ("expand").</summary>
        public sealed class SearchList : ComboBox
        {
            readonly ToolStripDropDown popup = new ToolStripDropDown { AutoClose = true, Padding = Padding.Empty };
            readonly ToolStripControlHost host;
            readonly Panel body = new Panel(), gap = new Panel { Dock = DockStyle.Top };
            public readonly TextBox Field = new TextBox { Dock = DockStyle.Top };                                  // test hooks
            public readonly ListBox Matches;
            public readonly bool Multi;
            public readonly HashSet<string> Checked = new HashSet<string>(); // Multi: the keys of the rows that are on
            public event EventHandler CheckedChanged;                         // Multi: a row turned on or off by the user
            bool syncing;
            public bool PopupOpen { get; private set; }
            public bool PopupVisible { get { return popup.Visible; } } // test hook
            public Control PopupBody { get { return body; } }
            int closedAt = Environment.TickCount - 10000;

            public SearchList() : this(false) { }

            public SearchList(bool multi)
            {
                Multi = multi;
                Matches = multi ? new CheckList { CheckOnClick = true } : new ListBox();
                Matches.Dock = DockStyle.Fill; Matches.IntegralHeight = false; Matches.BorderStyle = BorderStyle.None;
                body.Controls.Add(Matches); body.Controls.Add(gap); body.Controls.Add(Field);
                host = new ToolStripControlHost(body) { AutoSize = false, Margin = Padding.Empty, Padding = Padding.Empty };
                popup.Items.Add(host);
                popup.Closed += (s, e) => { PopupOpen = false; closedAt = Environment.TickCount; };
                Field.TextChanged += (s, e) => Refilter();
                // IsInputKey: these keys go to KeyDown, not to the dialog keys of the drop-down (ToolStripDropDown moves
                // between its items with them). In Outlook, WinForms sees the keys through its message hook while the
                // drop-down is open (ToolStripManager.ModalMenuFilter), so PreviewKeyDown also comes there.
                Field.PreviewKeyDown += (s, e) => { if (Nav(e.KeyCode)) e.IsInputKey = true; };
                Field.KeyDown += (s, e) => { if (Nav(e.KeyCode)) { e.Handled = e.SuppressKeyPress = true; FieldKey(e.KeyCode); } };
                if (multi) ((CheckedListBox)Matches).ItemCheck += (s, e) => OnItemCheck(e);
                else Matches.MouseClick += (s, e) => { int i = Matches.IndexFromPoint(e.Location); if (i >= 0) Take(Matches.Items[i]); };
            }

            /// <summary>Check boxes at the size of the row: CheckedListBox draws them at 13 px when the program has no
            /// high-DPI setting (outlook.exe.config has none), too small on a 2x screen (2.19 layout test, 10-23). A click
            /// anywhere on the row turns it on or off (CheckOnClick), so the hit area does not change.</summary>
            sealed class CheckList : CheckedListBox
            {
                protected override void OnDrawItem(DrawItemEventArgs e)
                {
                    if (e.Index < 0 || e.Index >= Items.Count) return;
                    e.DrawBackground();
                    int sz = Math.Max(11, e.Bounds.Height * 3 / 5), pad = Math.Max(2, sz / 4);
                    var box = new Rectangle(e.Bounds.X + pad, e.Bounds.Y + (e.Bounds.Height - sz) / 2, sz, sz);
                    if (Items[e.Index] is Choice) // not for the note "(eşleşen yok)"
                        ControlPaint.DrawCheckBox(e.Graphics, box, ButtonState.Flat | (GetItemChecked(e.Index) ? ButtonState.Checked : ButtonState.Normal));
                    bool sel = (e.State & DrawItemState.Selected) != 0;
                    TextRenderer.DrawText(e.Graphics, GetItemText(Items[e.Index]), e.Font, new Rectangle(box.Right + pad, e.Bounds.Y, e.Bounds.Right - box.Right - pad, e.Bounds.Height),
                                          sel ? SystemColors.HighlightText : ForeColor,
                                          TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
                    e.DrawFocusRectangle();
                }
            }

            /// <summary>Multi: a row turns on or off (click, Space, Enter). A note row and "(tümü)" alone do not turn off.</summary>
            void OnItemCheck(ItemCheckEventArgs e)
            {
                if (syncing) return;
                var c = Matches.Items[e.Index] as Choice;
                bool on = e.NewValue == CheckState.Checked;
                if (c == null || c == Choice.All && !on) { e.NewValue = e.CurrentValue; return; }
                if (c == Choice.All) Checked.Clear(); else if (on) Checked.Add(c.Key); else Checked.Remove(c.Key);
                SyncChecks(e.Index); // the other rows; this row takes e.NewValue
                ShowChecked();
                if (CheckedChanged != null) CheckedChanged(this, EventArgs.Empty);
            }

            void SyncChecks(int skip)
            {
                var clb = (CheckedListBox)Matches;
                syncing = true;
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    var c = clb.Items[i] as Choice;
                    if (i != skip && c != null) clb.SetItemChecked(i, c == Choice.All ? Checked.Count == 0 : Checked.Contains(c.Key));
                }
                syncing = false;
            }

            /// <summary>Multi: the texts of the rows that are on, in list order.</summary>
            public List<string> CheckedTexts() { return Items.Cast<object>().Skip(1).OfType<Choice>().Where(c => Checked.Contains(c.Key)).Select(c => c.Text).ToList(); }

            /// <summary>Multi: item 0 of the closed box shows the choice (it is always the selected item).</summary>
            public void ShowChecked()
            {
                var t = CheckedTexts();
                var top = t.Count == 0 ? Choice.All : new Choice { Text = t.Count == 1 ? t[0] : t.Count + " klasör seçili" };
                if (Items.Count == 0) Items.Add(top); else if (Items[0].ToString() != top.Text) Items[0] = top;
                SelectedIndex = 0;
            }

            /// <summary>Multi: set the rows that are on (test hook; "Filtreleri temizle" gives an empty list).</summary>
            public void SetChecked(IEnumerable<string> keys)
            {
                Checked.Clear(); Checked.UnionWith(keys);
                ShowChecked();
                if (CheckedChanged != null) CheckedChanged(this, EventArgs.Empty);
            }

            static bool Nav(Keys k) { return k == Keys.Up || k == Keys.Down || k == Keys.PageUp || k == Keys.PageDown || k == Keys.Enter || k == Keys.Escape; }

            /// <summary>Lower case without marks: "Çakır" -> "cakir", "İSTANBUL" -> "istanbul".</summary>
            public static string Fold(string s)
            {
                var sb = new StringBuilder(s.Length);
                foreach (char c in s.Normalize(NormalizationForm.FormD))
                    if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c == 'ı' ? 'i' : char.ToLowerInvariant(c));
                return sb.ToString();
            }

            void Refilter()
            {
                var words = Fold(Field.Text).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                // Item 0 of the box: "(tümü)" (Multi: the text of the choice); the list shows the real "(tümü)" row.
                var rows = Items.Cast<object>().Skip(1);
                var hits = (words.Length == 0 ? new object[] { Choice.All }.Concat(rows) : rows.Where(o => { string t = Fold(o.ToString()); return words.All(t.Contains); })).ToArray();
                Matches.BeginUpdate();
                Matches.Items.Clear();
                if (hits.Length > 0) Matches.Items.AddRange(hits); else Matches.Items.Add("(eşleşen yok)");
                if (Multi) SyncChecks(-1);
                Matches.SelectedIndex = hits.Length == 0 ? -1 : words.Length == 0 ? Math.Max(0, SelectedIndex) : 0;
                Matches.EndUpdate();
            }

            /// <summary>A key in the search field (also a test hook).</summary>
            public void FieldKey(Keys k)
            {
                int n = Matches.Items.Count, page = Math.Max(1, Matches.ClientSize.Height / Math.Max(1, Matches.ItemHeight) - 1);
                if (k == Keys.Escape) { popup.Close(); Focus(); }
                else if (k == Keys.Enter && Multi)
                {
                    var clb = (CheckedListBox)Matches; // turn the row on or off (OnItemCheck); the list stays open
                    if (clb.SelectedIndex >= 0) clb.SetItemChecked(clb.SelectedIndex, !clb.GetItemChecked(clb.SelectedIndex));
                }
                else if (k == Keys.Enter) { if (Matches.SelectedItem != null) Take(Matches.SelectedItem); }
                else if (Matches.SelectedIndex >= 0)
                    Matches.SelectedIndex = Math.Max(0, Math.Min(n - 1, Matches.SelectedIndex + (k == Keys.Down ? 1 : k == Keys.Up ? -1 : k == Keys.PageDown ? page : -page)));
            }

            void Take(object item)
            {
                int i = Items.IndexOf(item);
                if (i < 0) return; // "(eşleşen yok)"
                popup.Close();
                SelectedIndex = i; // SelectedIndexChanged: the filter changed
                Focus();
            }

            public void ClosePopup() { if (PopupOpen) popup.Close(); }

            /// <summary>Open the list with the search field (text: the first letters).</summary>
            public void ShowPopup(string text)
            {
                if (!Enabled || PopupOpen) return;
                Field.Font = Matches.Font = Font;
                Field.AccessibleName = AccessibleName + ": ara"; Matches.AccessibleName = AccessibleName;
                Field.AccessibleDescription = "Ad ya da adres yazın. Ok tuşlarıyla seçin, Enter ile alın, Esc ile kapatın.";
                Matches.AccessibleDescription = Multi ? "Enter ya da boşluk işareti değiştirir. Liste açık kalır." : "Enter seçer.";
                // ListBox.ItemHeight is a fixed 13 px before its window exists: the row height comes from the font.
                // CheckedListBox rows are Font.Height + 2 (the last row was cut in the 2.19 layout test, 10-23).
                int pad = Math.Max(2, Font.Height / 5), rows = Math.Min(MaxDropDownItems, Math.Max(3, Items.Count)),
                    rowH = Math.Max(TextRenderer.MeasureText("Ag", Font).Height + 1, Multi ? Matches.ItemHeight : 0);
                body.Padding = new Padding(pad); gap.Height = pad;
                Field.Text = text; Refilter();
                body.Size = host.Size = new Size(Math.Max(Width, DropDownWidth + (Multi ? rowH + pad : 0)), Field.PreferredHeight + pad * 3 + rows * rowH); // Multi: + check box
                popup.Show(this, new Point(0, Height));
                PopupOpen = true; // after Show: a Show that fails does not block Fill
                if (Environment.OSVersion.Platform == PlatformID.Win32NT) SendMessage(Field.Handle, 0x1501, (IntPtr)1, "Aramak için yazın"); // EM_SETCUEBANNER
                Field.Focus(); Field.SelectionStart = Field.TextLength;
            }

            [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);

            protected override void WndProc(ref Message m)
            {
                // A click opens the search list, not the native list. A click that closed the open list (outside click on
                // this box) does not open it again.
                if ((m.Msg == 0x201 || m.Msg == 0x203) && Enabled) // WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
                {
                    if (unchecked(Environment.TickCount - closedAt) > 300) { Focus(); ShowPopup(""); }
                    return;
                }
                if (m.Msg == 0x20A && Multi) return; // WM_MOUSEWHEEL: it would show one folder in the box, not choose it
                base.WndProc(ref m);
            }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                // Multi: the keys that move the selection of a closed list open it (item 0 is the choice, not a folder).
                bool move = Multi && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down || e.KeyCode == Keys.PageUp || e.KeyCode == Keys.PageDown ||
                                      e.KeyCode == Keys.Home || e.KeyCode == Keys.End);
                if (e.KeyCode == Keys.F4 && !e.Alt || (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up) && e.Alt || move) { ShowPopup(""); e.Handled = e.SuppressKeyPress = true; return; }
                base.OnKeyDown(e);
            }

            protected override void OnKeyPress(KeyPressEventArgs e)
            {
                if (!char.IsControl(e.KeyChar)) { ShowPopup(e.KeyChar.ToString()); e.Handled = true; return; }
                base.OnKeyPress(e);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) popup.Dispose();
                base.Dispose(disposing);
            }
        }

        /// <summary>Fill a filter list (sorted by the Turkish alphabet after "(tümü)") and keep the chosen value by its key; a
        /// value that is gone gives "(tümü)". An open list is not changed (the user is choosing).</summary>
        void Fill(ComboBox box, IEnumerable<Choice> values)
        {
            var sl = box as SearchList;
            if (box.DroppedDown || sl != null && sl.PopupOpen) return;
            string keep = KeyOf(box);
            var all = values.OrderBy(c => c.Text, StringComparer.Create(TR, true)).ToList();
            fillingFilters = true;
            box.BeginUpdate();
            box.Items.Clear(); box.Items.Add(Choice.All);
            box.Items.AddRange(all.ToArray());
            if (sl != null && sl.Multi) { sl.Checked.IntersectWith(all.Select(c => c.Key)); sl.ShowChecked(); } // the folders that are still there
            else box.SelectedIndex = (keep == null ? -1 : all.FindIndex(c => c.Key == keep)) + 1;
            box.EndUpdate();
            fillingFilters = false;
            int w = all.Select(c => TextRenderer.MeasureText(c.Text, box.Font).Width).DefaultIfEmpty(0).Max() + SystemInformation.VerticalScrollBarWidth + S(8);
            box.DropDownWidth = Math.Max(box.Width, Math.Min(S(500), w));
        }

        /// <summary>Sender and recipient values of the index file (done.facets): [[key, name, address, mails]], [[key, name, mails]].</summary>
        void ShowFacets(Dictionary<string, object> f)
        {
            Func<string, IEnumerable<ArrayList>> rows = k => f != null && f.ContainsKey(k) && f[k] is ArrayList ? ((ArrayList)f[k]).Cast<ArrayList>() : Enumerable.Empty<ArrayList>();
            Fill(senderBox, rows("senders").Select(r =>
            {
                string name = r[1] as string ?? "", addr = r[2] as string ?? "";
                return new Choice { Key = (string)r[0], Text = addr.Length == 0 || string.Equals(addr, name, StringComparison.OrdinalIgnoreCase) ? name : name + " <" + addr + ">" };
            }));
            Fill(recipientBox, rows("recipients").Select(r => new Choice { Key = (string)r[0], Text = r[1] as string ?? "" }));
        }

        /// <summary>Folder filter = one or more of the folders checked on the İndeks tab (each folder is one line there, so its
        /// subfolders are not included). Rebuilt only when the checked folders change.</summary>
        void RefreshFolderFilter()
        {
            var sel = SelectedFolders();
            var now = new HashSet<string>(folderBox.Items.Cast<Choice>().Skip(1).Select(c => c.Key)); // the list is sorted by name
            if (!now.SetEquals(sel.Select(f => f.Id))) Fill(folderBox, sel.Select(f => new Choice { Key = f.Id, Text = f.Path.TrimStart('\\') }));
        }

        /// <summary>The scope and the filters of a search: the folder filter narrows the checked folders to the chosen ones;
        /// the sender and recipient keys, the date range and the attachment filter go to the worker (begin.filters). note:
        /// the active filters for the status text ("" = none).</summary>
        public List<FolderInfo> FilteredScope(List<FolderInfo> sel, out Dictionary<string, object> filters, out string note)
        {
            if (folderBox.Checked.Count > 0) sel = sel.Where(f => folderBox.Checked.Contains(f.Id)).ToList();
            filters = new Dictionary<string, object> { {"sender", KeyOf(senderBox)}, {"recipient", KeyOf(recipientBox)}, {"attach", KeyOf(attFilterBox)} };
            DateTime? from, to;
            string dates = DateRange(out from, out to);
            if (from != null) filters["date_from"] = from.Value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            if (to != null) filters["date_to"] = to.Value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            var parts = new List<string>();
            foreach (var b in new[] { Tuple.Create("gönderen", senderBox), Tuple.Create("alıcı", recipientBox) })
                if (KeyOf(b.Item2) != null) parts.Add(b.Item1 + " " + b.Item2.Text);
            if (folderBox.Checked.Count > 0) parts.Add("klasör " + string.Join(", ", folderBox.CheckedTexts()));
            if (dates != null) parts.Add("tarih " + dates);
            if (KeyOf(attFilterBox) != null) parts.Add("ek " + attFilterBox.Text);
            note = parts.Count == 0 ? "" : " Filtre: " + string.Join(", ", parts) + ".";
            return sel;
        }

        /// <summary>The date filter as [from, to) in local time (null = open end) and its text for the status line (null = no
        /// date filter). The presets count from today at the search start. "Özel aralık…": both days included; a start after
        /// the end is read as the same range the other way round.</summary>
        string DateRange(out DateTime? from, out DateTime? to)
        {
            from = to = null;
            var t = DateTime.Today;
            switch (KeyOf(dateBox))
            {
                case "7d": from = t.AddDays(-6); break; // today and the 6 days before
                case "30d": from = t.AddDays(-29); break;
                case "3m": from = t.AddMonths(-3); break;
                case "12m": from = t.AddMonths(-12); break;
                case "year": from = new DateTime(t.Year, 1, 1); break;
                case "lastyear": from = new DateTime(t.Year - 1, 1, 1); to = new DateTime(t.Year, 1, 1); break;
                case "custom":
                    DateTime a = dateFrom.Value.Date, b = dateTo.Value.Date;
                    if (a > b) { var x = a; a = b; b = x; }
                    from = a; to = b.AddDays(1);
                    return a.ToString("dd.MM.yyyy", TR) + "–" + b.ToString("dd.MM.yyyy", TR);
                default: return null;
            }
            return dateBox.Text;
        }

        ComboBox[] AllFilterBoxes { get { return new[] { senderBox, recipientBox, folderBox, dateBox, attFilterBox }; } }

        /// <summary>Preview height: the user's height (previewWant) when the result list has room for its header and 3
        /// rows; else smaller, but not below 100 px (96 DPI) or the user's height.</summary>
        void FitPreview()
        {
            fitQueued = false;
            if (IsDisposed) return;
            int need = 4 * (TextRenderer.MeasureText("Ag", list.Font).Height + S(6)); // header + 3 rows
            int h = Math.Max(Math.Min(previewWant, S(100)), Math.Min(previewWant, previewBox.Height + list.Height - need));
            if (h != previewBox.Height) previewBox.Height = h;
        }

        /// <summary>A filter changed (by the user or by a refill): the date row shows only for "Özel aralık…", the
        /// clear button only with a filter. The search does not start again by itself.</summary>
        void FilterChanged()
        {
            FilterRows();
            clearFiltersBtn.Enabled = AllFilterBoxes.Any(b => b.SelectedIndex > 0) || folderBox.Checked.Count > 0;
            if (!fillingFilters && session == null) SetStatus("Filtre değişti. Uygulamak için Ara'ya basın.");
        }

        /// <summary>The third filter row (Başlangıç, Bitiş) shows only for "Özel aralık…".</summary>
        void FilterRows()
        {
            bool range = KeyOf(dateBox) == "custom";
            int rowH = (int)filterTable.RowStyles[0].Height;
            foreach (var c in rangeRow) c.Visible = range;
            filterTable.RowStyles[2].Height = range ? rowH : 0;
            filterTable.Height = filterTable.Padding.Vertical + (range ? 3 : 2) * rowH;
        }

        void ClearFilters()
        {
            fillingFilters = true;
            foreach (var b in AllFilterBoxes) b.SelectedIndex = 0;
            folderBox.SetChecked(new string[0]);
            fillingFilters = false;
            FilterChanged();
            senderBox.Focus();
        }

        // ------------------------------------------------------------------ result list: sort, shortcut menu
        /// <summary>A click on a column header: sort by it; a second click turns the order. First order: dates and Güven
        /// descending (newest, most relevant first), texts A to Z.</summary>
        void HeaderClick(int col) { SortResults(col, col == sortCol ? !sortDesc : col == 2 || col == 3); }

        void SortResults(int col, bool desc)
        {
            sortCol = col; sortDesc = desc;
            list.Sort();
            ShowSortArrow();
            if (list.SelectedItems.Count > 0) list.SelectedItems[0].EnsureVisible();
        }

        void ShowSortArrow()
        {
            for (int i = 0; i < ResultColumns.Length; i++)
                list.Columns[i].Text = ResultColumns[i] + (i == sortCol ? (sortDesc ? " ▼" : " ▲") : "");
        }

        /// <summary>Order of the result list. Güven = the worker's order (rank: relevance); Tarih = the mail date; the other
        /// columns by their text in the Turkish alphabet. Equal values keep the worker's order.</summary>
        sealed class ResultOrder : IComparer
        {
            static readonly StringComparer Text = StringComparer.Create(TR, true);
            readonly SearchForm f;
            public ResultOrder(SearchForm f) { this.f = f; }
            static int Rank(ListViewItem it) { object v; return ((Dictionary<string, object>)it.Tag).TryGetValue("rank", out v) ? (int)v : 0; }
            static DateTime When(ListViewItem it)
            {
                DateTime d;
                return DateTime.TryParse(((Dictionary<string, object>)it.Tag)["date"] as string, CultureInfo.InvariantCulture, DateTimeStyles.None, out d) ? d : DateTime.MinValue;
            }
            public int Compare(object x, object y)
            {
                ListViewItem a = (ListViewItem)x, b = (ListViewItem)y;
                int c = f.sortCol == 2 ? When(a).CompareTo(When(b))
                      : f.sortCol == 3 ? Rank(b).CompareTo(Rank(a)) // ascending = the least relevant first
                      : Text.Compare(Cell(a, f.sortCol), Cell(b, f.sortCol));
                if (f.sortDesc) c = -c;
                return c != 0 ? c : Rank(a).CompareTo(Rank(b));
            }
        }

        Dictionary<string, object> SelectedResult() { return list.SelectedItems.Count > 0 ? (Dictionary<string, object>)list.SelectedItems[0].Tag : null; }

        void CopySubject()
        {
            string s = Str(SelectedResult() ?? new Dictionary<string, object>(), "subject");
            if (string.IsNullOrEmpty(s)) return;
            try { Clipboard.SetText(s); SetStatus("Konu panoya kopyalandı."); }
            catch (ExternalException) { SetStatus("Konu kopyalanamadı. Pano başka bir programda açık olabilir; yeniden deneyin."); }
        }

        /// <summary>The Gönderen filter value of a result (results.sender_key), or null: an older worker, or a sender that is
        /// not in the list (the list has the 1,000 senders with the most mails).</summary>
        Choice SenderChoice(Dictionary<string, object> r)
        {
            string key = r != null ? Str(r, "sender_key") : null;
            return string.IsNullOrEmpty(key) ? null : senderBox.Items.OfType<Choice>().FirstOrDefault(c => c.Key == key);
        }

        /// <summary>"Yalnız bu gönderenle yeniden ara": the last search or "Benzerlerini bul" again (repeatLast), with the
        /// Gönderen filter of the result. Off while a search runs (it would not start) and before the first run.</summary>
        bool SearchRuns { get { return session != null && (session.Mode == "search" || session.Mode == "similar"); } }

        bool CanFilterBySender(Dictionary<string, object> r) { return SenderChoice(r) != null && repeatLast != null && !SearchRuns; }

        void FilterBySelectedSender()
        {
            if (SearchRuns) { SetStatus("Arama sürüyor. Bitince yeniden deneyin ya da Durdur'a basın."); return; }
            var c = SenderChoice(SelectedResult());
            if (c == null || repeatLast == null) return;
            senderBox.SelectedItem = c;
            repeatLast();
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
            if (session != null && session.Mode == "search") return;
            if (!ReadyOrQueue(() => { query.Text = q; StartSearch(); })) return;
            StoreItem st; List<FolderInfo> sel;
            if (!CanStart(out st, out sel)) return;
            Dictionary<string, object> filters;
            sel = FilteredScope(sel, out filters, out filterNote);
            RememberSearch(q);
            list.Items.Clear(); ShowPreview();
            EventNames.Clear(); modelSeen = false;
            lastStoreId = st.Id;
            Begin(new WorkerSession(this, source, settings, "search", st.Id, st.Name, sel, q) { Filters = filters });
            searchedText = query.Text; // "Arama bitince ilk sonuca geç": only while this text stays
            repeatLast = () => { query.Text = q; StartSearch(); };
            SetStatus(ModelState == "loading" ? "Arama başladı. Yapay zekâ modeli hazırlanıyor (pencere açılınca başladı)…"
                    : ModelState == "ready" ? "Arama başladı. Yapay zekâ hazır." : "Arama başladı…");
        }

        /// <summary>One worker at a time. No session: true. An index run: it stops after a confirmation (its work is saved)
        /// and retry runs after it. Another session (status check, Güncelle, a search): retry runs after it.</summary>
        bool ReadyOrQueue(Action retry)
        {
            if (session == null) return true;
            if (session.Mode == "index")
            {
                if (MessageBox.Show(this, "İndeksleme sürüyor. Arama için indeksleme durdurulsun mu?\n\nYapılan iş kaydedilir. " +
                                    "İndekslemeyi sonra İndeks sekmesinden kaldığı yerden sürdürebilirsiniz.", "HatırMail",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return false;
                queued = retry;
                StopSession();
                return false;
            }
            queued = retry;
            if (session.Mode == "update") SetStatus("Arama, güncelleme bitince başlayacak.");
            return false;
        }

        /// <summary>"Benzerlerini bul" (Outlook shortcut menu): the mails nearest to this mail, by its stored vectors (worker
        /// mode similar: no model load, no query). The Ara filters apply. folderId null: no exclusion check (tests).</summary>
        public void FindSimilar(string eid, string storeId, string folderId)
        {
            tabs.SelectedTab = tabSearch;
            var cur = storeBox.SelectedItem as StoreItem;
            if (cur == null || cur.Id != storeId)
            {
                if (Busy) { SetStatus("E-posta başka bir posta kutusunda. Önce çalışan aramayı ya da indekslemeyi durdurun."); return; }
                int i = storeBox.Items.Cast<StoreItem>().ToList().FindIndex(x => x.Id == storeId);
                if (i < 0) { SetStatus("Bu e-postanın posta kutusu bulunamadı."); return; }
                storeBox.SelectedIndex = i; // LoadFolders: the folder list and a status check; the search waits for it
            }
            if (folderId != null && excludedIds.Contains(folderId))
            {
                SetStatus("Bu e-postanın klasörü Ayarlar'da indeksten hariç tutulmuş. Benzerleri aranamaz.");
                return;
            }
            if (!ReadyOrQueue(() => FindSimilar(eid, storeId, folderId))) return;
            StoreItem st; List<FolderInfo> sel;
            if (!CanStart(out st, out sel)) return;
            Dictionary<string, object> filters;
            sel = FilteredScope(sel, out filters, out filterNote);
            list.Items.Clear(); ShowPreview();
            EventNames.Clear(); modelSeen = false;
            lastStoreId = st.Id;
            Begin(new WorkerSession(this, source, settings, "similar", st.Id, st.Name, sel, "") { Filters = filters, SimilarEid = eid });
            repeatLast = () => FindSimilar(eid, storeId, folderId);
            SetStatus("Benzer e-postalar aranıyor…");
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

        /// <summary>"Güncelle": only the mails that came, changed, moved or were deleted after the last index run (worker mode
        /// update). New mails of a folder that was never indexed completely wait for "İndekslemeyi başlat".</summary>
        public void StartUpdate()
        {
            if (session != null)
            {
                if (session.Mode == "status") queued = StartUpdate;
                return;
            }
            StoreItem st; List<FolderInfo> sel;
            if (!CanStart(out st, out sel)) return;
            EventNames.Clear(); modelSeen = false;
            Begin(new WorkerSession(this, source, settings, "update", st.Id, st.Name, sel, ""));
            phase.Text = "Güncelleme başladı: klasör listeleri okunuyor…";
            SetStatus("Güncelleme başladı: yalnız son indekslemeden sonraki değişiklikler işlenir.");
        }

        static bool IndexLike(string mode) { return mode == "index" || mode == "update"; } // no query, no time limit, Durdur on the İndeks tab
        static string ModeName(string mode) { return mode == "index" ? "İndeksleme" : mode == "update" ? "Güncelleme" : mode == "similar" ? "Benzer arama" : "Arama"; }

        /// <summary>Index state of the selected folders: a short worker run (no scan, no mail reading). With keep-alive on,
        /// the worker then loads the model in the background when the store has a semantic index (preload), so that the
        /// first search does not wait for it. Runs only on a user action: window open, store change, "Durumu yenile",
        /// after saving settings.</summary>
        public void StartStatus()
        {
            if (session != null) return;
            var st = storeBox.SelectedItem as StoreItem;
            // The scope can be empty (every folder excluded): the check then only deletes the excluded folders.
            if (st == null || allFolders.Count == 0 || !File.Exists(settings.Python) || !File.Exists(settings.Worker)) return;
            EventNames.Clear();
            Begin(new WorkerSession(this, source, settings, "status", st.Id, st.Name, folders, "", preload: true));
        }

        int KeepMs { get { return KeepAliveMsOverride > 0 ? KeepAliveMsOverride : settings.KeepAliveMin * 60000; } }

        // The worker's own limit for the wait between sessions. The window closes the worker first (keepTimer); this only
        // ends a worker whose window hangs. ponytail: 10 min margin; a sleep of the PC can still let the worker end first
        // (then the next search starts a new worker, or, in a rare race, reports "hata" once).
        int KeepS { get { return KeepMs <= 0 ? 0 : KeepMs / 1000 + 600; } }

        void Begin(WorkerSession s)
        {
            if (closing) return; // the window closes: no new worker
            s.Exclude = excludedIds.ToList(); // every session deletes them from the index (the store cannot change meanwhile)
            if (s.Mode != "status" || !purgeAsked) notice = null;
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
                SetStatus((notice != null ? notice + " " : "") +
                          (state == "loading" ? "Yapay zekâ modeli arka planda hazırlanıyor. Bu sırada arama yapabilirsiniz."
                        : state == "ready" ? "Yapay zekâ hazır: aramalar beklemeden çalışır." : "Yapay zekâ modeli yüklenemedi. Arama yine çalışır; yeniden deneyin."));
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
            searchBtn.Enabled = m != "search"; indexBtn.Enabled = updateBtn.Enabled = !Busy; refreshBtn.Enabled = m == null;
            stopBtn.Enabled = m == "search"; indexStopBtn.Enabled = IndexLike(m);
            storeBox.Enabled = !Busy;
            deleteBtn.Enabled = m == null; // no worker may hold the index files
            excludeBtn.Enabled = m == null && allFolders.Count > 0;
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
                    if (mode != "status" && mode != "similar") SetStatus("Klasörler taranıyor (silinen ve taşınan postalar)…");
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
                    LastDoneStats = Dict(ev, "stats");
                    if (ev.ContainsKey("facets") && ev["facets"] != null) ShowFacets(Dict(ev, "facets")); // older worker: lists stay
                    if (stt == "error" && mode != "status")
                        SetStatus(ModeName(mode) + " tamamlanamadı (" + (ev.ContainsKey("error") ? ev["error"] : "") + "). Yeniden deneyin.");
                    if (stt == "cancelled")
                        SetStatus("Durduruldu. Yapılan iş kaydedildi; sonraki arama ya da indeksleme kaldığı yerden sürer.");
                    if (stt == "ok" && mode == "index") SetStatus("İndeksleme bitti.");
                    if (stt == "ok" && mode == "update") SetStatus(UpdateText(LastDoneStats));
                    LastExcludedRemoved = ev.ContainsKey("excluded_removed") && ev["excluded_removed"] != null ? Convert.ToInt32(ev["excluded_removed"]) : -1;
                    if (mode == "status" && purgeAsked)
                    {
                        // The status check after "Klasörleri hariç tut": the model events after it keep this text in front.
                        purgeAsked = false;
                        notice = stt != "ok" ? "Klasör seçimi kaydedildi ama postalar indeksten silinemedi (" + (ev.ContainsKey("error") ? ev["error"] : "") +
                                               "). Durumu yenile'ye basarak yeniden deneyin."
                               : LastExcludedRemoved > 0 ? string.Format(TR, "Klasör seçimi kaydedildi. Hariç tutulan klasörlerden {0:N0} posta indeksten silindi.", LastExcludedRemoved)
                               : "Klasör seçimi kaydedildi. İndekste bu klasörlerden posta yoktu.";
                        SetStatus(notice);
                    }
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
            if (IndexLike(s.Mode))
                phase.Text = reason == "ok" ? ModeName(s.Mode) + " bitti: " + DateTime.Now.ToString("dd.MM.yyyy HH:mm", TR)
                           : (reason == "cancelled" ? ModeName(s.Mode) + " durduruldu." : ModeName(s.Mode) + " durdu (" + ReasonText(reason) + ").") +
                             (s.Mode == "update" ? " Kaldığı yerden sürdürmek için Güncelle'ye basın." : " Kaldığı yerden sürdürmek için İndekslemeyi başlat'a basın.");
            if (s.Mode != "status" && reason != "ok" && reason != "cancelled" && !status.Text.Contains("tamamlanamadı"))
                SetStatus(ModeName(s.Mode) + " durdu: " + ReasonText(reason) + ". Arama işlemi kapatıldı.");
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
            string eta = IndexLike(mode) && ev.ContainsKey("eta_s") && ev["eta_s"] != null ? " – bu aşamada kalan " + Duration(Convert.ToDouble(ev["eta_s"])) : "";
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
            if (IndexLike(mode)) phase.Text = text;
            SetStatus(text + (mode == "search" && ph != "scan" ? ". Durdurursanız kaldığı yerden sürer." : ""));
        }

        /// <summary>The result of "Güncelle" from done.stats: new and changed mails read, mails that left the index, new mails
        /// of never-indexed folders that were not read.</summary>
        static string UpdateText(Dictionary<string, object> st)
        {
            Func<string, int> n = k => st != null && st.ContainsKey(k) && st[k] != null ? Convert.ToInt32(st[k]) : 0;
            int added = n("new"), changed = n("changed") + n("same"), removed = n("removed"), skipped = n("unsynced_skipped");
            string text = added + changed + removed == 0 ? "İndeks güncel: son indekslemeden sonra yeni, değişen ya da silinen posta yok."
                : string.Format(TR, "Güncelleme bitti: {0:N0} yeni, {1:N0} değişen posta işlendi; {2:N0} posta indeksten çıktı (silinmiş ya da taşınmış).", added, changed, removed);
            if (skipped > 0)
                text += string.Format(TR, " {0:N0} posta hiç tam indekslenmemiş klasörlerde; onlar için İndekslemeyi başlat'a basın.", skipped);
            return text;
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
            // The next list of the same search keeps the mail that the user selected (not the first row that this method
            // selected: then the first row of the new list).
            var was = SelectedResult();
            string keep = was != null && (was["eid"] as string) != autoEid ? was["eid"] as string : null;
            list.BeginUpdate();
            list.Items.Clear();
            int rank = 0;
            foreach (var r in results)
            {
                r["rank"] = rank++; // the worker's order: the Güven column sorts by it, equal values keep it
                string date = r["date"] as string ?? "";
                DateTime dt;
                if (DateTime.TryParse(date, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt)) date = dt.ToString("dd.MM.yyyy HH:mm", TR);
                // The row text is one sentence for a screen reader (NVDA reads the text of each column, so the other
                // columns stay empty); DrawCell shows the columns from "cells".
                r["cells"] = new[] { r["subject"] as string ?? "", r["sender"] as string ?? "", date, Conf(r["confidence"] as string),
                                     string.Join(", ", ((ArrayList)r["why"]).Cast<string>()) };
                var it = new ListViewItem(new[] { RowText.Sentence(r, FolderPath(Str(r, "folder")), DateTime.Today), "", "", "", "" });
                it.Tag = r;
                if ((r["confidence"] as string) == "zayıf") it.ForeColor = SystemColors.GrayText;
                list.Items.Add(it);
            }
            list.Sort(); // the column that the user chose stays the order of new results
            list.EndUpdate();
            FitColumns(list); // a vertical scroll bar makes the list narrower
            var dw = ev.ContainsKey("date_windows") ? ((ArrayList)ev["date_windows"]).Cast<string>().ToList() : new List<string>();
            string dates = (dw.Count > 0 ? " Tarih aralığı: " + string.Join("; ", dw) + "." : "") + filterNote; // filterNote: the active filters
            bool semantic = ev.ContainsKey("semantic") && (bool)ev["semantic"];
            LastSemantic = semantic;
            if (ev.ContainsKey("similar_to")) // "Benzerlerini bul"
            {
                var to = Dict(ev, "similar_to");
                string subj = to != null ? "\"" + (to["subject"] as string) + "\"" : "";
                if (to == null)
                    SetStatus("Bu e-posta henüz indekste yok. İndeks sekmesinde Güncelle'ye basın, sonra yeniden deneyin.");
                else if (!true.Equals(to["embedded"]))
                    SetStatus("Bu e-postanın anlam indeksi henüz yok. İndeks sekmesinde Güncelle'ye basın, sonra yeniden deneyin.");
                else if (results.Count == 0)
                    SetStatus("Benzer e-posta bulunamadı: " + subj + "." + filterNote);
                else
                    SetStatus(results.Count + " benzer e-posta: " + subj + ". Gri satırlar yalnız zayıf benzerliktir." + filterNote);
            }
            else if (stage == "keyword")
                SetStatus((results.Count > 0 ? "Kelime eşleşmeleri gösteriliyor." : "Tam kelime eşleşmesi yok.") + " Anlamsal arama sürüyor…" + dates);
            else if (stage == "semantic") // the indexed mails; the new ones are indexed now, then the final list comes
                SetStatus((results.Count > 0 ? results.Count + " sonuç." : "Şimdilik sonuç yok.") +
                          " Yeni postalar ve ekler indeksleniyor. Bitince liste bir kez daha yenilenir." + dates);
            else if (!semantic)
                SetStatus((results.Count > 0 ? results.Count + " kelime eşleşmesi." : "Kelime eşleşmesi yok.") +
                          " Anlam indeksi henüz yok: İndeks sekmesinden indekslemeyi başlatın." + dates);
            else if (results.Count == 0 || results.All(r => (r["confidence"] as string) == "zayıf"))
                SetStatus("Güçlü bir eşleşme bulunamadı. Gri satırlar yalnız zayıf benzerliktir; başka ayrıntı deneyin." + dates);
            else
                SetStatus(results.Count + " sonuç. Açmak için çift tıklayın ya da Outlook'ta Aç'a basın." + dates);
            var sel = list.Items.Cast<ListViewItem>().FirstOrDefault(i => keep != null && (string)((Dictionary<string, object>)i.Tag)["eid"] == keep);
            if (sel != null) { sel.Selected = true; sel.EnsureVisible(); } // a kept row can be lower than the first page
            else if (list.Items.Count > 0) { list.Items[0].Selected = true; autoEid = (string)((Dictionary<string, object>)list.Items[0].Tag)["eid"]; } // as before 10-23
            else ShowPreview();
            UpdateButtons();
            // "Arama bitince ilk sonuca geç" (Ayarlar): the final list of a search, only while the user waits in the query
            // row with the same text (a user who went on to another field keeps the focus).
            if (stage == "final" && settings.FocusFirstResult && searchedText != null && list.Items.Count > 0 &&
                query.Text == searchedText && (ActiveControl == null || qrow.Contains(ActiveControl)))
            {
                var first = list.SelectedItems.Count > 0 ? list.SelectedItems[0] : list.Items[0];
                first.Selected = first.Focused = true;
                ActiveControl = list;
            }
            if (stage == "final" || ev.ContainsKey("similar_to")) searchedText = null;
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
            RefreshFolderFilter(); // the folder filter lists the checked folders
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
            else if (session != null && session.Mode == "update")
            {
                scopeLine.Text = "Güncelleme sürüyor: yeni ve değişen postalar işleniyor. Ara'ya basarsanız arama güncelleme bitince başlar.";
                summary.Text = "İndeks: anlam " + sem;
            }
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
            // No color carries the state (high contrast keeps the text color): an incomplete index is bold, as the full text says.
            bool warn = !(complete || cov.Count == 0 || Busy);
            if (warn) { if (!ReferenceEquals(scopeLine.Font, boldFont)) scopeLine.Font = boldFont; }
            else if (!ReferenceEquals(scopeLine.Font, scopeLine.Parent.Font)) scopeLine.ResetFont(); // the page font again
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
                attBar.Visible = false;
                return;
            }
            var r = (Dictionary<string, object>)list.SelectedItems[0].Tag;
            string inAtt = Str(r, "match_attachment"), member = Str(r, "match_att_member");
            attBar.Visible = !string.IsNullOrEmpty(inAtt);
            if (attBar.Visible)
            {
                attLine.Text = "Ekte bulundu: " + inAtt + (string.IsNullOrEmpty(member) ? "" : " › " + member);
                openAttBtn.AccessibleDescription = attLine.Text;
                openAttBtn.Enabled = AttIndex(r) > 0; // a result of an older worker has the name only
            }
            Add(r["subject"] as string ?? "", titleFont, SystemColors.WindowText, Marks(r, "hl_subject"));
            preview.AppendText("\n");
            string date = Cell(list.SelectedItems[0], 2);
            Add((r["sender"] as string ?? "") + (date.Length > 0 ? "  •  " + date : ""), Font, SystemColors.GrayText, null);
            string to = r.ContainsKey("recipients") ? r["recipients"] as string : "";
            if (!string.IsNullOrEmpty(to)) { preview.AppendText("\n"); Add("Kime: " + to, Font, SystemColors.GrayText, null); }
            preview.AppendText("\n\n");
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

        // Append text; marks = [start, length] in UTF-16 units of text (from the worker) are bold on the Info color (a
        // system color: light yellow in the default theme); high contrast: bold and underlined, the theme colors.
        void Add(string text, Font font, Color color, List<int[]> marks)
        {
            int pos = 0;
            foreach (var m in (marks ?? new List<int[]>()).Where(m => m[0] >= 0 && m[1] > 0).OrderBy(m => m[0]))
            {
                int a = Math.Max(pos, Math.Min(text.Length, m[0])), b = Math.Min(text.Length, m[0] + m[1]);
                if (b <= a) continue;
                Append(text.Substring(pos, a - pos), font, color, Color.Empty, false);
                Append(text.Substring(a, b - a), font, color, SystemColors.Info, true);
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
            bool hc = HC;
            Font bold;
            if (mark && !markFonts.TryGetValue(font, out bold))
                markFonts[font] = bold = new Font(font, hc ? FontStyle.Bold | FontStyle.Underline : FontStyle.Bold);
            preview.SelectionFont = mark ? markFonts[font] : font;
            preview.SelectionColor = hc ? SystemColors.WindowText : mark ? SystemColors.InfoText : color;
            preview.SelectionBackColor = mark && !hc ? back : SystemColors.Window;
            preview.Select(preview.TextLength, 0);
        }

        void OpenSelected()
        {
            if (list.SelectedItems.Count == 0) return;
            var r = (Dictionary<string, object>)list.SelectedItems[0].Tag;
            try { source.Open(r["eid"] as string, lastStoreId); }
            catch (Exception ex) { SetStatus("Posta açılamadı (taşınmış ya da silinmiş olabilir): " + ex.GetType().Name); }
        }

        static string Str(Dictionary<string, object> r, string key) { object v; return r.TryGetValue(key, out v) ? v as string : null; }
        static int AttIndex(Dictionary<string, object> r) { object v; return r.TryGetValue("match_att_index", out v) && v != null ? Convert.ToInt32(v) : 0; }

        /// <summary>"Eki aç": save a copy of the matched attachment (for a zip: the matched member only), protect it (read-only,
        /// Mark of the Web) and open it with its program. Outlook is not needed for the open; the mail is not changed.</summary>
        void OpenAttachment()
        {
            if (list.SelectedItems.Count == 0) return;
            var r = (Dictionary<string, object>)list.SelectedItems[0].Tag;
            string eid = r["eid"] as string, want = Str(r, "match_attachment"), member = Str(r, "match_att_member");
            int idx = AttIndex(r);
            if (idx < 1) return;
            string dir = null;
            try
            {
                bool item;
                string name = source.AttachmentName(eid, lastStoreId, idx, out item);
                // The index can be older than the mail: the attachment at this place must have the matched name.
                if (name == null || !string.Equals(name, want, StringComparison.OrdinalIgnoreCase))
                {
                    SetStatus("Ek bulunamadı. Posta değişmiş ya da silinmiş olabilir. Outlook'ta Aç ile postayı açın.");
                    return;
                }
                string file = AttachmentCopy.SafeFileName(name);
                if (item && !file.EndsWith(".msg", StringComparison.OrdinalIgnoreCase)) file += ".msg"; // SaveAsFile writes an Outlook item as .msg
                dir = Path.Combine(AttachmentCopy.Root(settings.Data), Guid.NewGuid().ToString("N").Substring(0, 8));
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, file);
                if (!source.SaveAttachment(eid, lastStoreId, idx, path)) { SetStatus("Ek bulunamadı. Outlook'ta Aç ile postayı açın."); AttachmentCopy.DeleteCopy(dir); return; }
                string shown = name, note = "";
                if (!string.IsNullOrEmpty(member) && file.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    string m = AttachmentCopy.ExtractMember(path, member);
                    if (m != null) { path = m; shown = Path.GetFileName(m); }
                    else note = " Eşleşen dosya zip içinde bulunamadı; zip açıldı.";
                }
                AttachmentCopy.Protect(path);
                Launch(path);
                SetStatus("Ek açıldı: " + shown + "." + note + " Bu bir kopyadır; değiştirirseniz Farklı Kaydet ile başka bir yere kaydedin. Kopya pencere kapanınca silinir.");
            }
            catch (Exception ex)
            {
                AttachmentCopy.DeleteCopy(dir);
                SetStatus("Ek açılamadı (" + ex.GetType().Name + "). Postayı Outlook'ta Aç ile açıp eki oradan açabilirsiniz.");
            }
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
            if (!allowClose && e.CloseReason == CloseReason.UserClosing && session != null && IndexLike(session.Mode) &&
                MessageBox.Show(this, ModeName(session.Mode) + " sürüyor. Pencere kapanırsa " + ModeName(session.Mode).ToLower(TR) + " durur (yapılan iş kaydedilir).\n\n" +
                                "Kapatılsın mı? İşin sürmesi için Hayır'a basın ve pencereyi simge durumuna küçültün.",
                                "HatırMail", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                e.Cancel = true;
                return;
            }
            closing = true; queued = null; // a search waiting for the end of an index run is dropped
            if (session != null) session.Stop("window closed");
            CloseWarm("window closed"); // the model never outlives the window
            AttachmentCopy.Clean(settings.Data); // opened attachment copies (a file that a program holds: next window open)
        }

        /// <summary>Outlook closes or unloads the add-in: close the window and stop every worker.</summary>
        public void Shutdown() { allowClose = true; Close(); }

        // Test hooks (TestHost)
        public void SetQuery(string q) { query.Text = q; }
        public int ResultCount { get { return list.Items.Count; } }
        public string ResultEid(int i) { return ((Dictionary<string, object>)list.Items[i].Tag)["eid"] as string; }
        public void SelectResult(int i) { list.Items[i].Selected = true; }
        public string SelectedEid { get { var r = SelectedResult(); return r != null ? r["eid"] as string : null; } } // test hook
        public void ClickOpen() { OpenSelected(); }
        public void ClickOpenAttachment() { OpenAttachment(); }
        public string AttachLine { get { return attBar.Visible ? attLine.Text : null; } } // null: no attachment bar
        public bool AttachButtonEnabled { get { return attBar.Visible && openAttBtn.Enabled; } }
        public void ClickStop() { StopSession(); }
        public string StatusText { get { return status.Text; } }
        public string CoverageText { get { return scopeLine.Text + " | " + pctRead.Text + " | " + pctSem.Text + " | " + pctAtt.Text; } }
        public bool IndexComplete { get { return summary.Text == "İndeks: tam"; } }
        public bool ModelLoaded { get { return modelSeen; } }
        public string SessionMode { get { return session != null ? session.Mode : null; } }
        public string LastPhase; public int LastDone, LastTotal; public bool LastSemantic; // test hooks: last progress / results
        public Dictionary<string, object> LastDoneStats;                                  // test hook: stats of the last done
        public ComboBox[] FilterBoxes { get { return AllFilterBoxes; } } // test hook: Gönderen, Alıcı, Klasör, Tarih, Ek
        public DateTimePicker[] DatePickers { get { return new[] { dateFrom, dateTo }; } }    // test hooks: filters, sorting, result menu
        public bool ClearFiltersEnabled { get { return clearFiltersBtn.Enabled; } }
        public void ClickClearFilters() { ClearFilters(); }
        public void ClickHeader(int col) { HeaderClick(col); }
        public string HeaderText(int col) { return list.Columns[col].Text; }
        public bool SelectedHasSenderChoice { get { return SenderChoice(SelectedResult()) != null; } }
        public bool FilterBySenderEnabled { get { return CanFilterBySender(SelectedResult()); } }
        public void ClickFilterBySender() { FilterBySelectedSender(); }
        public void SetFolderChecked(string id, bool on)                                 // test hook: a click on the İndeks list
        {
            foreach (ListViewItem it in folderList.Items) if (((FolderInfo)it.Tag).Id == id) it.Checked = on;
            UpdateSummary();
        }
        public string PhaseText { get { return phase.Text; } }                            // test hook
        public List<string> ResultWhy(int i) { return ((ArrayList)((Dictionary<string, object>)list.Items[i].Tag)["why"]).Cast<string>().ToList(); } // test hook
        public string ResultConfidence(int i) { return ((Dictionary<string, object>)list.Items[i].Tag)["confidence"] as string; }             // test hook
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
                for (int i = 0; i < preview.TextLength; i++) { preview.Select(i, 1); var f = preview.SelectionFont; if (f != null && f.Bold && (!HC || f.Underline)) n++; }
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
        public List<string> ListedFolderIds() { return folderList.Items.Cast<ListViewItem>().Select(i => ((FolderInfo)i.Tag).Id).ToList(); } // test hook
        public List<string> ExcludedFolderIds { get { return excludedIds.ToList(); } }                                                 // test hook
        public List<FolderInfo> AllFolders { get { return allFolders; } }                                                             // test hook
        public bool ExcludeButtonEnabled { get { return excludeBtn.Enabled; } }                                                       // test hook
        public void Feed(string mode, Dictionary<string, object> ev) { OnEvent(mode, ev); } // test hook: a worker event
        public string IndexFileOf(string storeId) { return IndexFile(storeId); }          // test hook
        public string QueryText { get { return query.Text; } }                            // test hooks: searches
        public SearchHistory History { get { return history; } }
        public void ReloadHistory() { history = SearchHistory.Load(settings.Data); RefreshSuggestions(); } // as at the window open
        public string[] Suggestions { get { return query.AutoCompleteCustomSource.Cast<string>().ToArray(); } }
        public List<ToolStripItem> HistoryMenuItems() { BuildHistoryMenu(); return historyMenu.Items.Cast<ToolStripItem>().ToList(); }

        // Test hooks: accessibility, keyboard, scale (TestHost --layout-check)
        public bool PressKey(Keys k) { var m = new Message(); return ProcessCmdKey(ref m, k); } // a shortcut, as ProcessCmdKey gets it
        public bool PressMnemonic(char c) { return ProcessMnemonic(c); }                           // Alt+letter
        public Control Active { get { return ActiveControl; } }
        public TabControl Tabs { get { return tabs; } }
        public ListView Results { get { return list; } }
        public Rectangle RingBounds { get { UpdateRing(); return ring.Visible ? ring.Bounds : Rectangle.Empty; } } // in root coordinates
        public void SimulateSystemColorsChanged() { OnSystemColorsChanged(EventArgs.Empty); }
        public string ResultSentence(int i) { return list.Items[i].Text; }
        public string[] ResultCells(int i) { return (string[])((Dictionary<string, object>)list.Items[i].Tag)["cells"]; }
        public void MarkSearchStarted() { searchedText = query.Text; }
        public int TextScalePct { get { return textPct; } }
        public int WindowDpi { get { return dpi; } }
        public Control QueryBox { get { return query; } }
        public ContextMenuStrip ResultMenu { get { return resultMenu; } }
        public void SetActive(Control c) { ActiveControl = c; }
        public bool FocusFirstSetting { get { return settings.FocusFirstResult; } set { settings.FocusFirstResult = value; } }

        class StoreItem { public string Id, Name; public override string ToString() { return Name; } }
    }

    /// <summary>The text of a result row: one sentence that a screen reader reads at once, for example "Ayşe Yılmaz,
    /// Bütçe taslağı, 12 Mart, Gelen Kutusu, ekte eşleşti: rapor.xlsx". Parts: sender, subject, date (with the year
    /// when it is not this year), the last part of the folder path, the attachment match, "zayıf eşleşme". An empty part
    /// is left out. No WinForms: TestHost --row-text-check runs it on mono.</summary>
    public static class RowText
    {
        static readonly CultureInfo TR = CultureInfo.GetCultureInfo("tr-TR");

        public static string Sentence(Dictionary<string, object> r, string folderPath, DateTime today)
        {
            Func<string, string> str = key => { object v; return r.TryGetValue(key, out v) ? (v as string ?? "").Trim() : ""; };
            var parts = new List<string> { str("sender"), str("subject") };
            DateTime d;
            if (DateTime.TryParse(str("date"), CultureInfo.InvariantCulture, DateTimeStyles.None, out d))
                parts.Add(d.ToString(d.Year == today.Year ? "d MMMM" : "d MMMM yyyy", TR));
            string folder = (folderPath ?? "").TrimEnd('\\');
            parts.Add(folder.Substring(folder.LastIndexOf('\\') + 1));
            string att = str("match_attachment"), member = str("match_att_member").Replace('\\', '/');
            if (att.Length > 0)
                parts.Add("ekte eşleşti: " + (member.Length > 0 ? member.Substring(member.LastIndexOf('/') + 1) + " (" + att + " içinde)" : att));
            if (str("confidence") == "zayıf") parts.Add("zayıf eşleşme");
            return string.Join(", ", parts.Where(x => x.Length > 0));
        }
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
                            settings.Model + "\" --threads " + Threads + " --memory-mb " + MemoryLimitMb,
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
        public List<string> Exclude = new List<string>(); // folder ids excluded in Ayarlar: the worker deletes them from the index
        public Dictionary<string, object> Filters;        // Ara filters (sender, recipient keys); null = none
        public string SimilarEid;                         // mode similar: the mail of "Benzerlerini bul"
        /// <summary>status and similar: no scan, no mail reading, no attachment folder (the worker only reads its index).</summary>
        public bool Light { get { return Mode == "status" || Mode == "similar"; } }

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
            if (settings.Attachments && !Light)
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
                {"attachments", Light ? settings.Attachments : attDir != null}, {"att_dir", attDir},
                {"store", new Dictionary<string, object> { {"id", storeId}, {"name", storeName} }},
                {"folders", scope.Select(f => new Dictionary<string, object> { {"id", f.Id}, {"path", f.Path} }).ToList()}, {"exclude", Exclude} };
            if (Mode == "search") begin["index_budget_s"] = settings.SearchIndexBudgetS;
            if (Filters != null) begin["filters"] = Filters;
            if (SimilarEid != null) begin["similar_eid"] = SimilarEid;
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
            Mark(name, ev);
            if (name == "hello" && !Light) SendManifests();
            else if (name == "need") SendItems(((ArrayList)ev["eids"]).Cast<string>().ToList());
            else if (name == "need_att") SaveAttachments((ArrayList)ev["atts"]);
            else if (name == "done")
            {
                string st = ev["status"] as string;
                End(st == "ok" ? "ok" : st == "cancelled" ? "cancelled" : "hata");
            }
        }

        // Timing log (search only): one JSON line per search in data\timing.log. Numbers only: no query, mail text, entry
        // id, folder name or path (the folder rows are in scope order). Outlook side: the manifest of each folder [rows, ms,
        // ms to the first row (GetTable)], the mail reading, the attachment saving. Worker side: done.timing and done.stats.
        // ms values count from the session start. ponytail: two files of at most 1 MB (timing.log, timing.log.1).
        readonly Dictionary<string, object> log = new Dictionary<string, object>();
        readonly List<object> logFolders = new List<object>();
        long attMs; int attSaved;
        static readonly string[] Marks = { "hello", "need", "keyword", "semantic", "final", "done" };

        void Mark(string name, Dictionary<string, object> ev)
        {
            string key = name == "results" ? ev["stage"] as string : name;
            if (Array.IndexOf(Marks, key) >= 0 && !log.ContainsKey(key + "_ms")) log[key + "_ms"] = clock.ElapsedMilliseconds;
            if (name == "need") log["need"] = ((ArrayList)ev["eids"]).Count;
            if (name == "results" && key == "final") log["results"] = ((ArrayList)ev["results"]).Count;
            if (name == "done") { log["worker_timing"] = Nums(ev.ContainsKey("timing") ? ev["timing"] : null); log["worker_stats"] = Nums(ev.ContainsKey("stats") ? ev["stats"] : null); }
        }

        /// <summary>Only the numbers and true/false values of a worker object (the keys are fixed names of the worker).</summary>
        static object Nums(object v)
        {
            var d = v as Dictionary<string, object>;
            if (d != null) return d.Where(kv => Nums(kv.Value) != null).ToDictionary(kv => kv.Key, kv => Nums(kv.Value));
            return v is bool || v is int || v is long || v is decimal || v is double ? v : null;
        }

        void WriteLog(string reason)
        {
            if (Mode != "search") return;
            try
            {
                log["at"] = DateTime.Now.ToString("yyyy-MM-dd'T'HH:mm:ss");
                log["reason"] = reason; // a fixed word of this class ("ok", "cancelled", "hata", "zaman aşımı" …)
                log["total_ms"] = clock.ElapsedMilliseconds;
                log["manifest"] = logFolders;
                log["att_save_ms"] = attMs; log["att_saved"] = attSaved;
                log["threads"] = settings.Threads; log["budget_s"] = settings.SearchIndexBudgetS;
                string path = Path.Combine(settings.Data, "timing.log");
                var fi = new FileInfo(path);
                if (fi.Exists && fi.Length > 1 << 20) { File.Delete(path + ".1"); File.Move(path, path + ".1"); }
                File.AppendAllText(path, new JavaScriptSerializer().Serialize(log) + "\n", new UTF8Encoding(false));
            }
            catch { } // extra information only: a locked file or a full disk never stops a search
        }

        // Folder item lists. Pumps UI messages between pages so that "Durdur" stays clickable.
        void SendManifests()
        {
            var map = new Dictionary<string, string>();
            foreach (var f in scope)
            {
                var page = new List<object>();
                var fw = Stopwatch.StartNew();
                long first = -1; int n = 0;
                foreach (var row in source.Manifest(f, () => cancelled))
                {
                    if (first < 0) first = fw.ElapsedMilliseconds;
                    n++;
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
                logFolders.Add(new long[] { n, fw.ElapsedMilliseconds, first < 0 ? fw.ElapsedMilliseconds : first });
            }
            folderOf = map;
            log["manifest_end_ms"] = clock.ElapsedMilliseconds;
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
            log["read_ms"] = sw.ElapsedMilliseconds; log["read_failed"] = failed.Count; log["read_stopped"] = stopped;
            Send(new Dictionary<string, object> { {"op", "items_done"}, {"failed", failed}, {"stopped", stopped} });
        }

        // Save the asked attachments ([eid, index, name]) into attDir. The file name is a counter, never the
        // attachment name (no path tricks, no long names). A file that cannot be saved is reported as null.
        void SaveAttachments(ArrayList atts)
        {
            var files = new List<object>();
            var aw = Stopwatch.StartNew();
            try { SaveEach(atts, files); }
            finally { attMs += aw.ElapsedMilliseconds; }
            if (cancelled || ended) return;
            attSaved += files.Count(f => ((object[])f)[2] != null);
            Send(new Dictionary<string, object> { {"op", "att_files"}, {"files", files} });
        }

        void SaveEach(ArrayList atts, List<object> files)
        {
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
            WriteLog(reason);
            form.OnSessionEnd(this, reason);
        }
    }

    /// <summary>Windows Job Object: the worker dies with the job handle (and with Outlook); memory cap. DPI helpers.</summary>
    public static class Native
    {
        [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")] static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hwnd);
        [DllImport("user32.dll")] static extern int GetAwarenessFromDpiAwarenessContext(IntPtr context);

        /// <summary>The DPI of a window, or 0 (Windows before 10 1607, or not Windows).</summary>
        public static int WindowDpi(IntPtr hwnd)
        {
            try { return (int)GetDpiForWindow(hwnd); } catch (EntryPointNotFoundException) { return 0; } catch (DllNotFoundException) { return 0; }
        }

        /// <summary>The DPI awareness of a window: 0 unaware, 1 system, 2 per monitor; -1 unknown (test hook).</summary>
        public static int WindowAwareness(IntPtr hwnd)
        {
            try { return GetAwarenessFromDpiAwarenessContext(GetWindowDpiAwarenessContext(hwnd)); }
            catch (EntryPointNotFoundException) { return -1; } catch (DllNotFoundException) { return -1; }
        }

        /// <summary>Windows that this thread makes inside the block are per-monitor DPI aware (v1, as Outlook itself): they
        /// get WM_DPICHANGED on a monitor with another DPI. Office calls the ribbon callbacks in a system-aware context
        /// (learn.microsoft.com, "Handle high DPI and DPI scaling in your Office solution"); the old context comes back at
        /// the end. No effect before Windows 10 1607. ponytail: v1 does not scale the title bar of the window on another
        /// monitor; upgrade path: v2 (-4), after a test of the WinForms 4.8 child windows with it.</summary>
        public sealed class DpiScope : IDisposable
        {
            readonly IntPtr old;
            public DpiScope()
            {
                if (Environment.OSVersion.Platform != PlatformID.Win32NT) return;
                try { old = SetThreadDpiAwarenessContext((IntPtr)(-3)); } catch (EntryPointNotFoundException) { }
            }
            public void Dispose()
            {
                if (old != IntPtr.Zero) try { SetThreadDpiAwarenessContext(old); } catch (EntryPointNotFoundException) { }
            }
        }

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

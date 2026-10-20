// TestHost: runs the real SearchForm + WorkerSession against a JSON corpus (JsonSource), without Outlook.
// Usage: HatirMail.TestHost.exe <corpus.json> <hatirmail_home> <query> [expected_eid] [options]
//          --index              run "İndekslemeyi başlat" first (full index), then the search
//          --stop-after-ms N    press Durdur N ms after the start (of the index run with --index, else of the search)
//          --budget S           indexing seconds in the search (settings.json value is not changed)
//          --expect-keyword     the search must end without the model (no semantic index yet)
//          --expect-incomplete  the index must stay incomplete after the search (its time limit stopped the work)
//          --read-delay-ms N    every mail read takes N ms more (a slow Outlook: tests the read time limit of the add-in)
//          --max-ms N           the search must end within N ms
//          --warm N             window open: wait for the model preload (status check), then N searches in the same worker;
//                               each one without a model load (keep-alive must be on)
//          --keep-alive-ms N    before the last search: keep-alive N ms; after it the worker must close by itself
//          --keep-alive-min N   keep-alive setting for this run (0 = old behavior: the worker exits after each session)
//        HatirMail.TestHost.exe <corpus.json> <hatirmail_home> --layout-check [screenshot_dir]
// Exit code 0 = pass. Prints only event names, counts, ids and timings (no mail text).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace HatirMail
{
    static class TestHost
    {
        [STAThread]
        static int Main(string[] args)
        {
            // UTF-8 output (no console in the test runner: the default would be the ANSI code page).
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true });
            // As in Outlook: common controls 6 (visual styles). Outlook showed a ListView error that TestHost without
            // visual styles did not show (16 Oct). An unhandled exception is a test failure, not a dialog.
            Application.EnableVisualStyles();
            Application.ThreadException += (s, e) => { Console.WriteLine("FAIL unhandled UI exception: " + e.Exception); Environment.Exit(70); };
            AppDomain.CurrentDomain.UnhandledException += (s, e) => { Console.WriteLine("FAIL unhandled exception: " + e.ExceptionObject); Environment.Exit(70); };
            if (args.Length < 3) { Console.Error.WriteLine("usage: TestHost <corpus.json> <home> <query> [expected_eid] [options]"); return 64; }
            Environment.SetEnvironmentVariable("HATIRMAIL_HOME", args[1]);
            var settings = Settings.Load();
            if (args[2] == "--layout-check") return LayoutCheck(new JsonSource(args[0]), settings, args.Length > 3 ? args[3] : null);
            string expected = args.Length > 3 && !args[3].StartsWith("--") ? args[3] : null;
            bool index = args.Contains("--index"), expectKeyword = args.Contains("--expect-keyword"), expectIncomplete = args.Contains("--expect-incomplete");
            int stopAfter = Opt(args, "--stop-after-ms", -1), budget = Opt(args, "--budget", -1), maxMs = Opt(args, "--max-ms", -1);
            int warmRuns = Opt(args, "--warm", 0), keepMs = Opt(args, "--keep-alive-ms", 0), keepMin = Opt(args, "--keep-alive-min", -1);
            if (budget >= 0) settings.SearchIndexBudgetS = budget;
            if (keepMin >= 0) settings.KeepAliveMin = keepMin;
            var src = new JsonSource(args[0]);
            var form = new SearchForm(new SlowSource(src, Opt(args, "--read-delay-ms", 0)), settings) { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };
            int code = 1, searches = 0;
            var pids = new List<int>();
            var sw = Stopwatch.StartNew();
            var total = Stopwatch.StartNew();
            bool started = false, warmOk = true;
            int keepPid = 0;
            string lastReason = null;
            Action startMain = null, finish = null;
            form.WorkerStarted += (mode, pid) => { pids.Add(pid); Console.WriteLine("worker_start mode=" + mode + " pid=" + pid + " ms=" + sw.ElapsedMilliseconds); };
            form.ModelStateChanged += state =>
            {
                Console.WriteLine("model_state=" + (state ?? "closed") + " ms_since_window=" + total.ElapsedMilliseconds);
                // --warm: the first search starts when the model preloaded after the status check is ready.
                if (warmRuns > 0 && !started && state == "ready") form.BeginInvoke(startMain);
                if (warmRuns > 0 && !started && state == "error") { Console.WriteLine("FAIL preload error"); form.Shutdown(); }
                if (keepPid != 0 && state == null)
                {
                    // Keep-alive over: the kept worker must be gone, without a window close.
                    bool gone = !Alive(keepPid);
                    Console.WriteLine("keep_alive_close ms=" + sw.ElapsedMilliseconds + " worker_gone=" + gone + " -> " + (gone && sw.ElapsedMilliseconds >= keepMs ? "ok" : "FAIL"));
                    warmOk &= gone && sw.ElapsedMilliseconds >= keepMs;
                    keepPid = 0;
                    form.BeginInvoke(finish);
                }
            };
            startMain = () =>
            {
                started = true;
                Console.WriteLine("addin_mem_idle_before=" + Mem());
                sw.Restart();
                if (index) form.StartIndex(); else { form.SetQuery(args[2]); form.StartSearch(); }
                if (form.SessionMode == null) { Console.WriteLine("FAIL: no worker started"); form.Shutdown(); return; }
                if (stopAfter >= 0)
                {
                    var t = new Timer { Interval = Math.Max(1, stopAfter) };
                    t.Tick += (s2, e2) => { t.Stop(); Console.WriteLine("stop_click_ms=" + sw.ElapsedMilliseconds + " mode=" + form.SessionMode); form.ClickStop(); };
                    t.Start();
                }
            };
            form.SessionFinished += (mode, reason) =>
            {
                Console.WriteLine("session_end mode=" + mode + " reason=" + reason + " ms=" + sw.ElapsedMilliseconds);
                Console.WriteLine("coverage=" + form.CoverageText.Replace("\n", " ") + " complete=" + form.IndexComplete);
                if (mode == "status")
                {
                    // The window starts one status check on load (index state: no scan, no mail reading). --warm waits
                    // for the model preload after it (ModelStateChanged "ready").
                    if (started) return;
                    if (reason != "ok") { Console.WriteLine("FAIL status " + reason); form.Shutdown(); }
                    else if (warmRuns == 0) startMain();
                    return;
                }
                if (warmRuns > 0 && mode == "search")
                {
                    // Warm search: the same worker as the status check, no model load in this session, semantic results.
                    bool ok = reason == "ok" && !form.ModelLoaded && form.LastSemantic && form.ResultCount > 0 && pids.Count == 1;
                    Console.WriteLine("warm_search n={0} ms={1} model_loaded={2} semantic={3} results={4} workers_started={5} -> {6}",
                                      searches + 1, sw.ElapsedMilliseconds, form.ModelLoaded, form.LastSemantic, form.ResultCount, pids.Count, ok ? "ok" : "FAIL");
                    warmOk &= ok;
                    lastReason = reason;
                    if (++searches < warmRuns) { form.BeginInvoke((Action)(() => { sw.Restart(); form.SetQuery(args[2]); form.StartSearch(); })); return; }
                    if (keepMs > 0) { keepPid = form.WarmPid; sw.Restart(); form.SetKeepAliveMs(keepMs); return; } // finish comes at the close
                    form.BeginInvoke(finish);
                    return;
                }
                Console.WriteLine("events=" + string.Join(",", form.EventNames.Distinct()) + " last_progress=" + form.LastPhase + " " + form.LastDone + "/" + form.LastTotal);
                if (mode == "index" && stopAfter < 0)
                {
                    // The full index is done: now the search, as the user does after "İndekslemeyi başlat".
                    if (reason != "ok" || !form.IndexComplete) { Console.WriteLine("FAIL index " + reason); form.Shutdown(); return; }
                    sw.Restart();
                    form.SetQuery(args[2]); form.StartSearch();
                    return;
                }
                bool complete = form.IndexComplete; // before Report closes the window
                long ms = sw.ElapsedMilliseconds;
                Report(form, src, settings, pids, reason, mode, stopAfter, expected, expectKeyword, expectIncomplete, ref code);
                if (expectIncomplete && complete) { Console.WriteLine("FAIL: index complete after a time-limited search"); code = 1; }
                if (maxMs >= 0 && ms > maxMs) { Console.WriteLine("FAIL: search took " + ms + " ms, limit " + maxMs); code = 1; }
            };
            finish = () =>
            {
                Report(form, src, settings, pids, lastReason, "search", -1, expected, false, false, ref code);
                Console.WriteLine("warm_check=" + (warmOk ? "ok" : "FAIL") + " total_ms=" + total.ElapsedMilliseconds);
                if (!warmOk) code = 1;
            };
            form.Shown += (s, e) =>
            {
                // No store/folder list on an empty home: start at once (no status check runs then).
                if (!File.Exists(settings.Python)) startMain();
            };
            Application.Run(form);
            return code;
        }

        static bool Alive(int pid)
        {
            try { return pid > 0 && !Process.GetProcessById(pid).HasExited; } catch (ArgumentException) { return false; } catch { return true; }
        }

        static void Report(SearchForm form, JsonSource src, Settings settings, List<int> pids, string reason, string mode, int stopAfter,
                           string expected, bool expectKeyword, bool expectIncomplete, ref int code)
        {
            Console.WriteLine("events=" + string.Join(",", form.EventNames.Distinct()));
            Console.WriteLine("results=" + form.ResultCount + " semantic=" + form.LastSemantic + " model_loaded=" + form.ModelLoaded);
            for (int i = 0; i < Math.Min(3, form.ResultCount); i++)
                Console.WriteLine("result" + (i + 1) + "=" + form.ResultEid(i) + (form.ResultAttachment(i) != null ? " in_attachment" : ""));
            Console.WriteLine("attachments_saved=" + src.Saved);
            bool attLeft = Directory.Exists(Path.Combine(settings.Data, "tmp-att"));
            Console.WriteLine("att_tmp_left=" + attLeft);
            Console.WriteLine("worker_peak_rss_mb=" + form.WorkerPeakMb);
            Console.WriteLine("addin_mem_after=" + Mem());
            int marked = 0;
            if (form.ResultCount > 0) { form.SelectResult(0); marked = form.PreviewMarkedChars; form.ClickOpen(); }
            Console.WriteLine("preview_marked_chars=" + marked);
            Console.WriteLine("opened=" + string.Join(",", src.Opened));
            // The add-in finds the index file of the worker (same file name rule): the window starts with that store.
            bool indexFound = File.Exists(form.IndexFileOf(src.Stores()[0].Key));
            Console.WriteLine("index_file_found=" + indexFound);
            bool ok = !attLeft && pids.Count > 0 && indexFound;
            if (stopAfter >= 0) ok &= reason == "cancelled";
            else
            {
                ok &= reason == "ok" && mode == "search" && form.ResultCount > 0 && (expected == null || form.ResultEid(0) == expected) && src.Opened.Count == 1;
                // A time-limited search on a partial index is semantic only when some mails have vectors already.
                if (expectKeyword) ok &= !form.ModelLoaded && !form.LastSemantic;
                else if (!expectIncomplete) ok &= form.LastSemantic;
                ok &= marked > 0; // the preview marks the query words
            }
            // Keep-alive: the worker can stay while the window is open. The window closes now: no worker may stay.
            int keptBefore = pids.Count(Alive);
            var sw = Stopwatch.StartNew();
            form.Shutdown();
            int alive = pids.Count(Alive);
            Console.WriteLine("workers_started=" + pids.Count + " alive_before_close=" + keptBefore + " alive_after_close=" + alive +
                              " close_ms=" + sw.ElapsedMilliseconds);
            ok &= alive == 0;
            code = ok ? 0 : 1;
            Console.WriteLine(ok ? "PASS" : "FAIL");
        }

        /// <summary>A mail source with a slow Read (a slow Outlook or a large mailbox): every other call is passed on.</summary>
        class SlowSource : IMailSource
        {
            readonly IMailSource s; readonly int delay;
            public SlowSource(IMailSource s, int delay) { this.s = s; this.delay = delay; }
            public List<KeyValuePair<string, string>> Stores() { return s.Stores(); }
            public List<FolderInfo> Folders(string storeId) { return s.Folders(storeId); }
            public IEnumerable<object[]> Manifest(FolderInfo f, Func<bool> cancelled) { return s.Manifest(f, cancelled); }
            public Dictionary<string, object> Read(string eid, string storeId, string folderId)
            {
                if (delay > 0) System.Threading.Thread.Sleep(delay);
                return s.Read(eid, storeId, folderId);
            }
            public bool SaveAttachment(string eid, string storeId, int idx, string path) { return s.SaveAttachment(eid, storeId, idx, path); }
            public void Open(string eid, string storeId) { s.Open(eid, storeId); }
        }

        static int Opt(string[] args, string name, int dflt)
        {
            int i = Array.IndexOf(args, name);
            return i > 0 && i + 1 < args.Length ? int.Parse(args[i + 1]) : dflt;
        }

        // ------------------------------------------------------------------ high-DPI layout check + screenshots
        // Scale 2 = a 192-DPI screen, as on 2.19 over RDP, simulated in the 96-DPI SSH session (1024x768): the window
        // is laid out at its default client size (and at its minimum size) inside an off-screen panel, so the
        // check does not depend on the size of this screen. The screenshots are TestHost renders, not Outlook.
        static int LayoutCheck(IMailSource src, Settings settings, string shotDir)
        {
            SearchForm.TestScale = 2F;
            var form = new SearchForm(src, settings) { StartPosition = FormStartPosition.Manual, Location = new Point(-32000, -32000), ShowInTaskbar = false };
            int code = 1;
            form.Shown += (s, e) =>
            {
                bool ok = true;
                // Fake worker events with realistic sizes (no worker in this check).
                form.Feed("index", Coverage(form));
                form.Feed("index", new Dictionary<string, object> { {"ev", "progress"}, {"phase", "embed"}, {"done", 1234}, {"total", 4321}, {"rate", 0.6}, {"eta_s", 5148} });
                form.Feed("search", Results());
                form.SetQuery("geçen ay Ahmet'ten gelen kargo faturası");
                // The window is limited to this 1024x768 screen; its root panel is not (a child window): lay it out at
                // the client size of the default window and of the minimum window at 192 DPI.
                var root = form.Root;
                root.Dock = DockStyle.None;
                foreach (var size in new[] { form.DesignClientSize, form.MinClientSize })
                {
                    root.Bounds = new Rectangle(Point.Empty, size);
                    for (int tab = 0; tab < 3; tab++)
                    {
                        form.ShowTab(tab); root.PerformLayout(); Application.DoEvents();
                        string name = (size == form.DesignClientSize ? "default" : "min") + "-" + new[] { "ara", "indeks", "ayarlar" }[tab];
                        ok &= Check(root, name);
                        if (shotDir != null) Shot(root, Path.Combine(shotDir, "ui-" + name + ".png"));
                    }
                }
                form.ShowTab(2);
                ok &= SettingsRoundTrip(form, settings);
                form.ShowTab(0);
                Console.WriteLine("client_sizes default={0} min={1} root={2}", form.DesignClientSize, form.MinClientSize, root.Size);
                Console.WriteLine("preview_marked_chars=" + form.PreviewMarkedChars);
                ok &= form.PreviewMarkedChars > 0;
                Console.WriteLine(ok ? "PASS" : "FAIL");
                code = ok ? 0 : 1;
                form.Shutdown();
            };
            Application.Run(form);
            return code;
        }

        // Ayarlar tab: set the fields, press Kaydet, read settings.json back. The first save creates the file, the second
        // replaces it. Then a hand-edited file with values out of range. The file is deleted at the end (stable screenshots).
        static bool SettingsRoundTrip(SearchForm form, Settings settings)
        {
            Directory.CreateDirectory(settings.Home);
            if (File.Exists(settings.ConfigFile)) File.Delete(settings.ConfigFile);
            var all = SearchForm.AllControls(form).ToList();
            Func<string, NumericUpDown> num = n => all.OfType<NumericUpDown>().First(x => x.AccessibleName == n);
            var att = all.OfType<CheckBox>().First(x => x.Text.StartsWith("Ek dosyalarının"));
            var save = all.OfType<Button>().First(x => x.Text == "&Kaydet");
            bool ok = true;
            foreach (var v in new[] { new { T = 6, M = 6144, B = 45, A = false, K = 30 }, new { T = 2, M = 3072, B = 0, A = true, K = 0 } })
            {
                num("İşlemci çekirdeği").Value = v.T; num("Bellek sınırı (MB)").Value = v.M; num("Aramada indeksleme (sn)").Value = v.B;
                num("Yapay zekâyı hazır tut (dk)").Value = v.K;
                att.Checked = v.A;
                save.PerformClick();
                var r = Settings.Load();
                bool same = r.Threads == v.T && r.MemoryLimitMb == v.M && r.SearchIndexBudgetS == v.B && r.Attachments == v.A &&
                            r.KeepAliveMin == v.K && form.StatusText.StartsWith("Ayarlar kaydedildi");
                Console.WriteLine("settings_save threads={0} memory={1} budget={2} attachments={3} keep_alive_min={4} -> {5}", r.Threads, r.MemoryLimitMb,
                                  r.SearchIndexBudgetS, r.Attachments, r.KeepAliveMin, same ? "ok" : "FAIL");
                ok &= same;
            }
            File.WriteAllText(settings.ConfigFile, "{\"threads\": 99, \"memory_limit_mb\": 10, \"search_index_budget_s\": -5, \"embed_budget_s\": 7, \"keep_alive_min\": 999}");
            var c = Settings.Load();
            bool clamp = c.Threads == 16 && c.MemoryLimitMb == 2048 && c.SearchIndexBudgetS == 0 && c.KeepAliveMin == 120;
            Console.WriteLine("settings_clamp threads={0} memory={1} budget={2} keep_alive_min={3} -> {4}", c.Threads, c.MemoryLimitMb, c.SearchIndexBudgetS,
                              c.KeepAliveMin, clamp ? "ok" : "FAIL");
            // A file without the new key (an older install): keep-alive is on with the default.
            File.WriteAllText(settings.ConfigFile, "{\"threads\": 4}");
            bool dflt = Settings.Load().KeepAliveMin == 10;
            Console.WriteLine("settings_keep_alive_default=" + (dflt ? "ok" : "FAIL"));
            clamp &= dflt;
            bool tmpLeft = File.Exists(settings.ConfigFile + ".tmp");
            Console.WriteLine("settings_tmp_left=" + tmpLeft);
            File.Delete(settings.ConfigFile);
            return ok && clamp && !tmpLeft;
        }

        // Checks on the visible controls: text fits (labels, buttons, check boxes, tabs), controls stay inside the
        // window, list columns fill the list without a cut.
        static bool Check(Control root, string name)
        {
            bool ok = true;
            var rr = root.RectangleToScreen(root.ClientRectangle);
            int n = 0;
            foreach (var c in SearchForm.AllControls(root).Where(c => c.Visible))
            {
                n++;
                // Inside the client area of its parent (and so inside the window): nothing is clipped.
                var r = c.RectangleToScreen(new Rectangle(Point.Empty, c.Size));
                var pr = c.Parent.RectangleToScreen(c.Parent.ClientRectangle);
                if (r.Width > 0 && r.Height > 0 && (!pr.Contains(r) || !rr.Contains(r)))
                { Console.WriteLine("{0}: outside {1} '{2}' {3} parent={4}", name, c.GetType().Name, Short(c.Text), r, pr); ok = false; }
                var l = c as Label;
                if (l != null && l.Text.Length > 0)
                {
                    // Height for the text wrapped at the label width (one line for a label that ends with "…"),
                    // and the longest word fits the width.
                    int w = l.ClientSize.Width - l.Padding.Horizontal;
                    int need = (l.AutoEllipsis ? TextRenderer.MeasureText("Ag", l.Font).Height
                                : TextRenderer.MeasureText(l.Text, l.Font, new Size(Math.Max(1, w), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height)
                               + l.Padding.Vertical;
                    int word = TextRenderer.MeasureText(LongestWord(l.Text), l.Font).Width;
                    if (l.Height < need - 2 || !l.AutoEllipsis && word > w + 2)
                    { Console.WriteLine("{0}: label cut '{1}' size={2} need_h={3} word_w={4}", name, Short(l.Text), l.Size, need, word); ok = false; }
                }
                var b = c as ButtonBase;
                if (b != null && b.Text.Length > 0)
                {
                    var need = TextRenderer.MeasureText(b.Text.Replace("&", ""), b.Font);
                    int extra = b is CheckBox ? SystemInformation.MenuCheckSize.Width : 8;
                    if (b.Width < need.Width + extra || b.Height < need.Height)
                    { Console.WriteLine("{0}: button cut '{1}' size={2} need={3}", name, Short(b.Text), b.Size, need); ok = false; }
                }
                // Edit fields: the text fits the height; the search box is wide; a number box shows its largest value.
                if (c is TextBoxBase && !(c is RichTextBox) || c is ComboBox)
                {
                    // The edit inside a NumericUpDown has no border: the whole control (with its border) holds the text.
                    int fh = TextRenderer.MeasureText("Ag", c.Font).Height, h = c.Parent is UpDownBase ? c.Parent.Height : c.Height;
                    bool fit = h >= fh + 2;
                    if (c.AccessibleName == "Arama metni") fit &= c.Width >= root.Width * 2 / 5;
                    var nud = c.Parent as NumericUpDown;
                    if (nud != null) fit &= c.Width >= TextRenderer.MeasureText(nud.Maximum.ToString("0"), c.Font).Width + 4;
                    if (c is ComboBox) fit &= c.Text.Length > 0 && c.Width >= TextRenderer.MeasureText(c.Text, c.Font).Width + c.Height;
                    if (!fit) { Console.WriteLine("{0}: edit field too small {1} '{2}' size={3} h={4} font_h={5}", name, c.GetType().Name, Short(c.Text), c.Size, h, fh); ok = false; }
                }
                var lv = c as ListView;
                if (lv != null && lv.View == View.Details)
                {
                    int sum = lv.Columns.Cast<ColumnHeader>().Sum(h => h.Width);
                    if (sum > lv.ClientSize.Width || sum < lv.ClientSize.Width - 8 || lv.Columns[0].Width < lv.Columns[1].Width)
                    { Console.WriteLine("{0}: columns {1} list {2} first {3}", name, sum, lv.ClientSize.Width, lv.Columns[0].Width); ok = false; }
                    // Every cell except the first column (subject, folder: long texts end with "…") fits its column.
                    foreach (ListViewItem it in lv.Items)
                        for (int col = 1; col < lv.Columns.Count; col++)
                            if (TextRenderer.MeasureText(it.SubItems[col].Text, lv.Font).Width + 12 > lv.Columns[col].Width)
                            { Console.WriteLine("{0}: cell cut col={1} '{2}' w={3}", name, col, it.SubItems[col].Text, lv.Columns[col].Width); ok = false; }
                }
                var tc = c as TabControl;
                if (tc != null)
                    for (int i = 0; i < tc.TabCount; i++)
                        if (tc.GetTabRect(i).Width < TextRenderer.MeasureText(tc.TabPages[i].Text, tc.Font).Width + 4)
                        { Console.WriteLine("{0}: tab cut {1}", name, i); ok = false; }
            }
            Console.WriteLine("{0}: controls={1} {2}", name, n, ok ? "ok" : "FAIL");
            return ok;
        }

        static bool IsShown(Control c)
        {
            for (var p = c; p != null; p = p.Parent) if (p is TabPage && ((TabControl)p.Parent).SelectedTab != p) return false;
            return true;
        }

        static string LongestWord(string t) { return t.Split(new[] { ' ', '\n' }, StringSplitOptions.RemoveEmptyEntries).OrderByDescending(w => w.Length).FirstOrDefault() ?? ""; }
        static string Short(string t) { return t.Length > 40 ? t.Substring(0, 40) + "…" : t; }

        // ------------------------------------------------------------------ screenshots (TestHost render)
        static void Shot(Control root, string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            using (var bmp = new Bitmap(root.Width, root.Height))
            {
                root.DrawToBitmap(bmp, new Rectangle(Point.Empty, root.Size));
                // DrawToBitmap leaves out some native parts in this off-screen session: the RichTextBox (no WM_PRINT
                // support), EDIT controls (TextBox, the number of a NumericUpDown) and the text of a ComboBox.
                // ponytail: TestHost paints them at their real bounds with their real font, colors and text.
                using (var g = Graphics.FromImage(bmp))
                    foreach (var c in SearchForm.AllControls(root).Where(c => c.Visible))
                    {
                        var r = root.RectangleToClient(c.RectangleToScreen(new Rectangle(Point.Empty, c.Size)));
                        if (c is RichTextBox) { g.FillRectangle(SystemBrushes.Window, r); PrintRtf((RichTextBox)c, g, r); }
                        else if (c is TextBox)
                        {
                            var tb = (TextBox)c;
                            g.FillRectangle(SystemBrushes.Window, r);
                            if (!(c.Parent is UpDownBase)) ControlPaint.DrawBorder3D(g, r, Border3DStyle.Sunken);
                            TextRenderer.DrawText(g, tb.Text, tb.Font, Rectangle.Inflate(r, -4, 0), tb.Enabled ? tb.ForeColor : SystemColors.GrayText,
                                                  TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis |
                                                  (tb.TextAlign == HorizontalAlignment.Right ? TextFormatFlags.Right : TextFormatFlags.Left));
                        }
                        else if (c is ComboBox)
                        {
                            var t = new Rectangle(r.X + 3, r.Y + 3, r.Width - r.Height - 6, r.Height - 6); // left of the drop-down button
                            g.FillRectangle(SystemBrushes.Window, t);
                            TextRenderer.DrawText(g, c.Text, c.Font, t, c.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
                        }
                    }
                bmp.Save(path, ImageFormat.Png);
            }
            Console.WriteLine("screenshot " + path);
        }

        [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] struct CHARRANGE { public int cpMin, cpMax; }
        [StructLayout(LayoutKind.Sequential)] struct FORMATRANGE { public IntPtr hdc, hdcTarget; public RECT rc, rcPage; public CHARRANGE chrg; }
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, ref FORMATRANGE l);
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

        static void PrintRtf(RichTextBox rtb, Graphics g, Rectangle r)
        {
            const int EM_FORMATRANGE = 0x439;
            float tw = 1440f / g.DpiX, th = 1440f / g.DpiY; // pixels -> twips
            var rc = new RECT { Left = (int)((r.Left + 4) * tw), Top = (int)((r.Top + 2) * th), Right = (int)((r.Right - 4) * tw), Bottom = (int)((r.Bottom - 2) * th) };
            IntPtr hdc = g.GetHdc();
            try
            {
                var fr = new FORMATRANGE { hdc = hdc, hdcTarget = hdc, rc = rc, rcPage = rc, chrg = new CHARRANGE { cpMin = 0, cpMax = -1 } };
                SendMessage(rtb.Handle, EM_FORMATRANGE, (IntPtr)1, ref fr);
                SendMessage(rtb.Handle, EM_FORMATRANGE, IntPtr.Zero, IntPtr.Zero); // free the cached format data
            }
            finally { g.ReleaseHdc(hdc); }
        }

        // Fake worker events (layout check only): the live store size, Turkish folder names, a long subject.
        static Dictionary<string, object> Coverage(SearchForm form)
        {
            var folders = new ArrayList();
            foreach (var f in form.FolderIds())
                folders.Add(new Dictionary<string, object> { {"id", f.Key}, {"path", f.Value}, {"total", 1200}, {"indexed", 1200}, {"read", 1180},
                                                             {"embedded", 410}, {"att_total", 262}, {"att_done", 95}, {"scanned_at", 1792134000.0} });
            return new Dictionary<string, object> { {"ev", "progress"}, {"phase", "embed"}, {"done", 0}, {"total", 1}, {"eta_s", null},
                                                    {"coverage", new Dictionary<string, object> { {"folders", folders} }} };
        }

        static Dictionary<string, object> Results()
        {
            Func<string, string, string, string, string, ArrayList, ArrayList, Dictionary<string, object>> r = (eid, subj, who, conf, snip, hl, why) =>
                new Dictionary<string, object> { {"eid", eid}, {"subject", subj}, {"sender", who}, {"recipients", "Furkan Çakır; Muhasebe Ekibi"},
                    {"date", "2026-09-12T09:41:00"}, {"confidence", conf}, {"why", why}, {"snippet", snip}, {"hl", hl},
                    {"hl_subject", new ArrayList { new ArrayList { 0, 5 }, new ArrayList { 6, 8 } }}, {"attachments", "Fatura_2026-09.pdf; İrsaliye.xlsx"},
                    {"match_attachment", null} };
            string s1 = "Merhaba Furkan Bey, eylül ayına ait kargo faturası ekte gönderilmiştir. Toplam tutar 4.250 TL olup ödeme vadesi 30 gündür. İyi çalışmalar.";
            var list = new ArrayList {
                r("E1", "Kargo faturası – Eylül 2026 (Hızlı Lojistik A.Ş.)", "Ahmet Demir", "yüksek", s1,
                  new ArrayList { new ArrayList { s1.IndexOf("kargo"), 5 }, new ArrayList { s1.IndexOf("faturası"), 8 } }, new ArrayList { "kelime 2/2", "anlam 0.83", "tarih uygun" }),
                r("E2", "Kargo faturası – Ağustos", "Ahmet Demir", "orta", "Ağustos kargo faturası ve irsaliye ekte.", new ArrayList { new ArrayList { 8, 5 }, new ArrayList { 14, 8 } },
                  new ArrayList { "kelime 2/2", "anlam 0.74", "tarih yakın" }),
                r("E3", "Depo sayım sonuçları", "Ayşe Kaya", "zayıf", "Sayım dosyası ektedir.", new ArrayList(), new ArrayList { "kelime 1/2", "anlam 0.61", "ekte", "tarih dışı" }) };
            return new Dictionary<string, object> { {"ev", "results"}, {"stage", "final"}, {"results", list}, {"semantic", true},
                                                    {"date_windows", new ArrayList { "01.09.2026 – 30.09.2026" }} };
        }

        static string Mem()
        {
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            var p = Process.GetCurrentProcess(); p.Refresh();
            return string.Format("ws_mb={0:0.0} private_mb={1:0.0} managed_mb={2:0.0}", p.WorkingSet64 / 1048576.0,
                                 p.PrivateMemorySize64 / 1048576.0, GC.GetTotalMemory(true) / 1048576.0);
        }
    }
}

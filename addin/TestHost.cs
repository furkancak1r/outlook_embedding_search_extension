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
//          --open-att SHA256    press "Eki aç" on the first result: the opened copy must have this SHA-256 (a zip member:
//                               the member bytes), the Mark of the Web (Windows) and be gone after the window close
//          --exclude FOLDER_ID  the folder is excluded in Ayarlar (this run only; settings.json does not change): the status
//                               check at the window open must delete its mails from the index and the list must not show it
//          --expect-removed N   with --exclude: the number of deleted mails (default: more than 0)
//          --update EID         with --index: the mail EID is not in Outlook during the full index (it comes later). Then
//                               "Güncelle" must read only this mail (1 new, in less time than the full index), a second
//                               "Güncelle" must find nothing ("İndeks güncel"), and the search must find EID
//          --expect-no-model    the search must give semantic final results without a model load (stored query vector)
//          --park MS            window open (model preload), then the user closes it: the window hides and keeps the same
//                               worker; open again: no new worker, the model is ready, a search loads no model; close
//                               again: the worker and the window must be gone after MS (the time after the close)
//          --battery-pause MS   with --index: MS after the start the PC goes on battery: the index run stops (work kept)
//                               and waits; then AC power: it starts again and completes; the index run is at Idle priority
//        HatirMail.TestHost.exe <corpus.json> <hatirmail_home> --layout-check [screenshot_dir]
//        HatirMail.TestHost.exe - <hatirmail_home> --history-check   (recent/saved searches file only, no window: runs on mono)
//        HatirMail.TestHost.exe <att_corpus.json> <hatirmail_home> --attach-check   ("Eki aç" file rules, no window: runs on mono)
//        HatirMail.TestHost.exe - <hatirmail_home> --exclude-check   ("Klasörleri hariç tut" rules, no window: runs on mono)
//        HatirMail.TestHost.exe - <hatirmail_home> --ribbon-check    ("Benzerlerini bul" menu XML and mail reference, no window: runs on mono)
//        HatirMail.TestHost.exe - <hatirmail_home> --row-text-check  (the one-sentence text of a result row, no window: runs on mono)
//        HatirMail.TestHost.exe <corpus.json> <hatirmail_home> --similar EID EXPECTED_EID
//                             full index, then "Benzerlerini bul" on EID: no model load, EXPECTED_EID first and "yüksek", EID not
//                             in the results; then the same with the Gönderen filter of EID (only that sender)
// Exit code 0 = pass. Prints only event names, counts, ids and timings (no mail text).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace HatirMail
{
    // Stand-ins for the Outlook Selection, mail and folder (--ribbon-check). Public top-level types: the late binding in
    // HatirMail.dll (Connect.MailRef) sees only public members of public types of another assembly.
    public class FakeFolder { public string EntryID { get; set; } public string StoreID { get; set; } }
    public class FakeMail { public string EntryID { get; set; } public FakeFolder Parent { get; set; } }
    public class FakeSelection
    {
        readonly object[] items;
        public FakeSelection(params object[] items) { this.items = items; }
        public int Count { get { return items.Length; } }
        public object this[int i] { get { return items[i - 1]; } } // 1-based, as Outlook
    }

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
            if (args[2] == "--history-check") return HistoryCheck(settings) ? 0 : 1;
            if (args[2] == "--attach-check") return AttachCheck(args[0], settings) ? 0 : 1;
            if (args[2] == "--exclude-check") return ExcludeCheck(settings) ? 0 : 1;
            if (args[2] == "--ribbon-check") return RibbonCheck() ? 0 : 1;
            if (args[2] == "--row-text-check") return RowTextCheck() ? 0 : 1;
            if (args[2] == "--similar") return SimilarRun(new JsonSource(args[0]), settings, args[3], args[4]);
            if (args[2] == "--layout-check") return LayoutCheck(new JsonSource(args[0]), settings, args.Length > 3 ? args[3] : null);
            if (args.Contains("--park")) return ParkRun(new JsonSource(args[0]), settings, args[2], Opt(args, "--park", 5000));
            string expected = args.Length > 3 && !args[3].StartsWith("--") ? args[3] : null;
            bool index = args.Contains("--index"), expectKeyword = args.Contains("--expect-keyword"), expectIncomplete = args.Contains("--expect-incomplete");
            bool expectNoModel = args.Contains("--expect-no-model");
            int batteryMs = Opt(args, "--battery-pause", -1);
            bool battery = false, batteryOk = batteryMs < 0;
            int indexRuns = 0;
            int stopAfter = Opt(args, "--stop-after-ms", -1), budget = Opt(args, "--budget", -1), maxMs = Opt(args, "--max-ms", -1);
            int warmRuns = Opt(args, "--warm", 0), keepMs = Opt(args, "--keep-alive-ms", 0), keepMin = Opt(args, "--keep-alive-min", -1);
            string openAtt = Array.IndexOf(args, "--open-att") > 0 ? args[Array.IndexOf(args, "--open-att") + 1] : null;
            if (budget >= 0) settings.SearchIndexBudgetS = budget;
            if (keepMin >= 0) settings.KeepAliveMin = keepMin;
            var src = new JsonSource(args[0]);
            string excludeId = Array.IndexOf(args, "--exclude") > 0 ? args[Array.IndexOf(args, "--exclude") + 1] : null;
            int expectRemoved = Opt(args, "--expect-removed", -1);
            bool excludeOk = excludeId == null;
            if (excludeId != null) settings.ExcludedFolders[src.Stores()[0].Key] = new List<string> { excludeId }; // in memory only
            var slow = new SlowSource(src, Opt(args, "--read-delay-ms", 0));
            string updateEid = Array.IndexOf(args, "--update") > 0 ? args[Array.IndexOf(args, "--update") + 1] : null;
            if (updateEid != null) slow.Hidden.Add(updateEid); // the mail comes after the full index
            long indexMs = 0, updateMs = 0;
            int updates = 0;
            bool updateOk = updateEid == null;
            SearchForm form = null;
            Func<string, int> stat = k => form.LastDoneStats != null && form.LastDoneStats.ContainsKey(k) ? Convert.ToInt32(form.LastDoneStats[k]) : -1;
            form = new SearchForm(slow, settings) { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };
            if (batteryMs >= 0) { settings.PauseOnBattery = true; form.OnBattery = () => battery; form.PowerCheckMs = 500; }
            var launched = new List<string>();
            form.Launch = p => launched.Add(p); // no program starts in the test: record the path
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
                // Priority: an index run at Idle, a search at BelowNormal (Windows; mono gives other names or none).
                Console.WriteLine("worker_priority mode=" + form.SessionMode + " priority=" + form.WorkerPriority);
                if (batteryMs >= 0)
                {
                    var t = new Timer { Interval = Math.Max(1, batteryMs) };
                    t.Tick += (s2, e2) => { t.Stop(); battery = true; Console.WriteLine("battery_on ms=" + sw.ElapsedMilliseconds + " mode=" + form.SessionMode); };
                    t.Start();
                }
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
                    if (excludeId != null && reason == "ok")
                    {
                        // The status check at the window open deleted the excluded folder from the index (begin.exclude).
                        bool listed = form.ListedFolderIds().Contains(excludeId);
                        excludeOk = (expectRemoved >= 0 ? form.LastExcludedRemoved == expectRemoved : form.LastExcludedRemoved > 0) &&
                                    !listed && form.ExcludedFolderIds.Contains(excludeId);
                        Console.WriteLine("exclude_purge folder={0} removed={1} listed={2} -> {3}", excludeId, form.LastExcludedRemoved, listed, excludeOk ? "ok" : "FAIL");
                    }
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
                if (mode == "index" && batteryMs >= 0 && ++indexRuns == 1)
                {
                    // On battery: the run stopped (its work is kept) and waits for AC power; it must not start on battery.
                    bool paused = reason == "cancelled" && form.PausedMode == "index" && !form.IndexComplete && form.PhaseText.Contains("pille");
                    Console.WriteLine("battery_pause reason={0} waiting={1} complete={2} -> {3}", reason, form.PausedMode, form.IndexComplete, paused ? "ok" : "FAIL");
                    batteryOk = paused;
                    if (!paused) { form.Shutdown(); return; }
                    var t = new Timer { Interval = 2000 };
                    t.Tick += (s2, e2) =>
                    {
                        t.Stop();
                        bool still = form.SessionMode == null && form.PausedMode == "index";
                        Console.WriteLine("battery_wait no_start={0} -> {1}", still, still ? "ok" : "FAIL");
                        batteryOk &= still;
                        battery = false; sw.Restart(); // AC power: the index run starts again by itself
                    };
                    t.Start();
                    return;
                }
                if (mode == "index" && batteryMs >= 0)
                {
                    bool resumed = reason == "ok" && form.IndexComplete && form.PausedMode == null;
                    Console.WriteLine("battery_resume reason={0} complete={1} ms={2} -> {3}", reason, form.IndexComplete, sw.ElapsedMilliseconds, resumed ? "ok" : "FAIL");
                    batteryOk &= resumed;
                }
                if (mode == "index" && stopAfter < 0)
                {
                    // The full index is done: now the search, as the user does after "İndekslemeyi başlat".
                    if (reason != "ok" || !form.IndexComplete) { Console.WriteLine("FAIL index " + reason); form.Shutdown(); return; }
                    indexMs = sw.ElapsedMilliseconds;
                    sw.Restart();
                    if (updateEid != null) { slow.Hidden.Clear(); form.StartUpdate(); return; } // the mail comes in
                    form.SetQuery(args[2]); form.StartSearch();
                    Console.WriteLine("worker_priority mode=" + form.SessionMode + " priority=" + form.WorkerPriority); // the kept worker of the index run
                    return;
                }
                if (mode == "update")
                {
                    // 1st "Güncelle": the new mail only. 2nd: nothing ("İndeks güncel"). Then the search.
                    bool ok = reason == "ok" && form.IndexComplete && stat("unsynced_skipped") == 0 && stat("removed") == 0 &&
                              (++updates == 1 ? stat("need") == 1 && stat("new") == 1 && form.StatusText.StartsWith("Güncelleme bitti: 1 yeni")
                                              : stat("need") == 0 && form.StatusText.StartsWith("İndeks güncel") && !form.ModelLoaded);
                    if (updates == 1) updateMs = sw.ElapsedMilliseconds;
                    Console.WriteLine("update n={0} ms={1} need={2} new={3} skipped={4} complete={5} phase_ok={6} -> {7}", updates, sw.ElapsedMilliseconds,
                                      stat("need"), stat("new"), stat("unsynced_skipped"), form.IndexComplete, form.PhaseText.StartsWith("Güncelleme bitti"), ok ? "ok" : "FAIL");
                    updateOk = (updates == 1 || updateOk) && ok && form.PhaseText.StartsWith("Güncelleme bitti");
                    sw.Restart();
                    if (updates == 1) { form.BeginInvoke((Action)form.StartUpdate); return; }
                    form.BeginInvoke((Action)(() => { form.SetQuery(args[2]); form.StartSearch(); }));
                    return;
                }
                bool complete = form.IndexComplete; // before Report closes the window
                bool modelLoadedAtEnd = form.ModelLoaded, semanticAtEnd = form.LastSemantic;
                long ms = sw.ElapsedMilliseconds;
                string opened = null;
                if (openAtt != null && reason == "ok" && form.ResultCount > 0)
                {
                    form.SelectResult(0);
                    form.ClickOpenAttachment();
                    opened = launched.LastOrDefault();
                    bool same = opened != null && File.Exists(opened) && Sha256(opened) == openAtt.ToLowerInvariant();
                    bool inRoot = opened != null && Path.GetFullPath(opened).StartsWith(Path.GetFullPath(AttachmentCopy.Root(settings.Data)) + Path.DirectorySeparatorChar);
                    bool ro = opened != null && File.Exists(opened) && (File.GetAttributes(opened) & FileAttributes.ReadOnly) != 0;
                    string zone = opened != null ? ZoneOf(opened) : null;
                    bool motw = Environment.OSVersion.Platform != PlatformID.Win32NT || zone != null && zone.Contains("ZoneId=3");
                    bool ok = same && inRoot && ro && motw && form.StatusText.StartsWith("Ek açıldı");
                    Console.WriteLine("open_att line={0} file={1} sha_ok={2} in_open_att={3} read_only={4} motw={5} -> {6}", form.AttachLine != null,
                                      opened != null ? Path.GetFileName(opened) : "-", same, inRoot, ro, motw, ok ? "ok" : "FAIL");
                    if (!ok) openAtt = "FAIL";
                }
                else if (openAtt != null) { Console.WriteLine("open_att no result -> FAIL"); openAtt = "FAIL"; }
                Report(form, src, settings, pids, reason, mode, stopAfter, expected, expectKeyword, expectIncomplete, ref code);
                if (expectIncomplete && complete) { Console.WriteLine("FAIL: index complete after a time-limited search"); code = 1; }
                if (maxMs >= 0 && ms > maxMs) { Console.WriteLine("FAIL: search took " + ms + " ms, limit " + maxMs); code = 1; }
                if (!excludeOk) { Console.WriteLine("FAIL: excluded folder"); code = 1; }
                if (!batteryOk) { Console.WriteLine("FAIL: battery pause"); code = 1; }
                if (expectNoModel)
                {
                    // A recent search (this query ran before in this data folder): the stored query vector, no model.
                    bool cached = form.LastDoneTiming != null && true.Equals(form.LastDoneTiming.ContainsKey("query_cached") ? form.LastDoneTiming["query_cached"] : null);
                    bool nm = reason == "ok" && !modelLoadedAtEnd && semanticAtEnd && cached;
                    Console.WriteLine("query_vector_cache model_loaded={0} semantic={1} query_cached={2} ms={3} -> {4}", modelLoadedAtEnd, semanticAtEnd, cached, ms, nm ? "ok" : "FAIL");
                    if (!nm) { Console.WriteLine("FAIL: stored query vector"); code = 1; }
                }
                if (updateEid != null)
                {
                    // Güncelle read one mail: it must take less time than the full index (both with the model load).
                    bool fast = updates == 2 && updateMs < indexMs;
                    Console.WriteLine("update_check full_index_ms={0} update_ms={1} updates={2} -> {3}", indexMs, updateMs, updates, updateOk && fast ? "ok" : "FAIL");
                    if (!updateOk || !fast) { Console.WriteLine("FAIL: Güncelle"); code = 1; }
                }
                if (openAtt != null)
                {
                    // Report closed the window: the copy is gone.
                    bool gone = opened != null && !File.Exists(opened) && !Directory.Exists(AttachmentCopy.Root(settings.Data));
                    Console.WriteLine("open_att_cleaned=" + (gone ? "ok" : "FAIL"));
                    if (openAtt == "FAIL" || !gone) { Console.WriteLine("FAIL: Eki aç"); code = 1; }
                }
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

        // Pencere kapandıktan sonra modeli tut (feature 1), the real window and worker: open (status check + model preload),
        // close by the user -> hidden, the same worker alive; open again (Show, as Connect.ShowForm) -> status check on the
        // same worker (no new worker), model ready, a search without a model load; close again -> after parkMs the worker
        // and the window are gone. Needs a full index with vectors in the data folder (run --index first).
        static int ParkRun(JsonSource src, Settings settings, string q, int parkMs)
        {
            var form = new SearchForm(src, settings) { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };
            form.KeepAfterCloseMsOverride = parkMs;
            var pids = new List<int>();
            int code = 1, step = 0, pid0 = 0;
            bool ok = true;
            var sw = Stopwatch.StartNew();
            Action<string, bool> check = (n, pass) => { Console.WriteLine("park_" + n + "=" + (pass ? "ok" : "FAIL") + " ms=" + sw.ElapsedMilliseconds); ok &= pass; };
            form.WorkerStarted += (mode, pid) => { pids.Add(pid); Console.WriteLine("worker_start mode=" + mode + " pid=" + pid); };
            form.ModelStateChanged += state =>
            {
                if (state != "ready") return;
                if (step == 0)
                {
                    // Window open, model ready: the user closes the window.
                    step = 1; pid0 = form.WarmPid;
                    form.BeginInvoke((Action)(() =>
                    {
                        form.Close();
                        check("hidden_worker_kept", !form.IsDisposed && !form.Visible && form.IsParked && pid0 != 0 && form.WarmPid == pid0 && Alive(pid0));
                        step = 2; sw.Restart();
                        form.Show(); // the ribbon button again
                    }));
                }
                else if (step == 2)
                {
                    // The status check of the new open ran on the kept worker: the model is ready at once.
                    step = 3;
                    check("reopen_same_worker", pids.Count == 1 && form.WarmPid == pid0 && !form.IsParked && form.Visible);
                    form.BeginInvoke((Action)(() => { sw.Restart(); form.SetQuery(q); form.StartSearch(); }));
                }
            };
            form.SessionFinished += (mode, reason) =>
            {
                Console.WriteLine("session_end mode=" + mode + " reason=" + reason + " ms=" + sw.ElapsedMilliseconds);
                if (reason != "ok") { check("session_" + mode, false); form.Shutdown(); return; }
                if (mode != "search" || step != 3) return;
                check("search_no_model_load", !form.ModelLoaded && form.LastSemantic && form.ResultCount > 0 && pids.Count == 1);
                step = 4;
                form.BeginInvoke((Action)(() =>
                {
                    sw.Restart();
                    form.Close(); // hidden again; after parkMs the worker and the window end
                    check("hidden_again", form.IsParked && Alive(pid0));
                }));
            };
            form.FormClosed += (s, e) =>
            {
                bool gone = !Alive(pid0);
                check("closed_after_time", step == 4 && gone && sw.ElapsedMilliseconds >= parkMs);
                code = ok ? 0 : 1;
                Console.WriteLine(ok ? "PASS" : "FAIL");
            };
            var guard = new Timer { Interval = 300000 };
            guard.Tick += (s, e) => { guard.Stop(); Console.WriteLine("park timeout step=" + step); ok = false; form.Shutdown(); };
            guard.Start();
            Application.Run(form);
            return code;
        }

        // "Benzerlerini bul" without Outlook: the ribbon XML (well-formed, the menu of a mail, the callback exists) and
        // Connect.MailRef with stand-ins for a Selection and a mail (late binding through dynamic, as with Outlook).
        static bool RibbonCheck()
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("ribbon_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            try
            {
                var x = System.Xml.Linq.XDocument.Parse(Connect.CustomUI);
                var ns = x.Root.Name.Namespace;
                var menu = x.Descendants(ns + "contextMenu").SingleOrDefault();
                var btn = menu == null ? null : menu.Elements(ns + "button").SingleOrDefault();
                check("xml", menu != null && (string)menu.Attribute("idMso") == "ContextMenuMailItem" && btn != null &&
                             (string)btn.Attribute("label") == "Benzerlerini bul" && (string)btn.Attribute("onAction") == "OnSimilar" &&
                             typeof(Connect).GetMethod("OnSimilar") != null && typeof(Connect).GetMethod("OnSearch") != null &&
                             new Connect().GetCustomUI("Microsoft.Outlook.Explorer") == Connect.CustomUI && new Connect().GetCustomUI("Microsoft.Outlook.Mail.Read") == null);
                var mail = new FakeMail { EntryID = "E1", Parent = new FakeFolder { EntryID = "F1", StoreID = "S1" } };
                string eid, sid, fid;
                Connect.MailRef(new FakeSelection(mail), out eid, out sid, out fid);
                bool sel = eid == "E1" && sid == "S1" && fid == "F1";
                Connect.MailRef(mail, out eid, out sid, out fid);
                check("mail_ref", sel && eid == "E1" && sid == "S1" && fid == "F1");
                bool threw = false;
                try { Connect.MailRef(new object(), out eid, out sid, out fid); } catch { threw = true; }
                check("no_mail", threw);
                // Folder list (Table.GetArray): the documented (column, row) order and (row, column), also for 4 x 4.
                var d = new DateTime(2026, 10, 23, 9, 5, 7);
                Func<int, object[]> row = i => new object[] { "E" + i, d.AddMinutes(i), "IPM.Note", i == 1 ? (object)null : d };
                foreach (int n in new[] { 1, 3, 4, 500 })
                {
                    Array cr = Array.CreateInstance(typeof(object), 4, n), rc = Array.CreateInstance(typeof(object), n, 4);
                    for (int r = 0; r < n; r++) for (int c = 0; c < 4; c++) { cr.SetValue(row(r)[c], c, r); rc.SetValue(row(r)[c], r, c); }
                    Func<List<object[]>, bool> same = got => got.Count == n && Enumerable.Range(0, n).All(r =>
                        (string)got[r][0] == "E" + r && (string)got[r][1] == d.AddMinutes(r).ToString("yyyy-MM-dd'T'HH:mm:ss") &&
                        (string)got[r][2] == "IPM.Note" && (string)got[r][3] == (r == 1 ? "" : "2026-10-23T09:05:07"));
                    check("manifest_rows_" + n, same(OutlookSource.ManifestRows(cr)) && same(OutlookSource.ManifestRows(rc)));
                }
            }
            catch (Exception ex) { Console.WriteLine("ribbon_check exception " + ex.GetType().Name + ": " + ex.Message); ok = false; }
            Console.WriteLine(ok ? "PASS" : "FAIL");
            return ok;
        }

        // "Benzerlerini bul" with the real window and worker (JsonSource, no Outlook): full index, then FindSimilar (the code
        // behind the menu command) on eid; then with the Gönderen filter of eid. No model load in the similar sessions.
        static int SimilarRun(JsonSource src, Settings settings, string eid, string expected)
        {
            var form = new SearchForm(src, settings) { ShowInTaskbar = false, WindowState = FormWindowState.Minimized };
            int code = 1, step = 0;
            string storeId = src.Stores()[0].Key, sender = null;
            var sw = Stopwatch.StartNew();
            form.SessionFinished += (mode, reason) =>
            {
                Console.WriteLine("session_end mode=" + mode + " reason=" + reason + " ms=" + sw.ElapsedMilliseconds);
                if (reason != "ok") { Console.WriteLine("FAIL " + mode + " " + reason); form.Shutdown(); return; }
                if (mode == "status" && step == 0) { step = 1; form.StartIndex(); return; }
                if (mode == "index") { step = 2; sw.Restart(); form.BeginInvoke((Action)(() => form.FindSimilar(eid, storeId, null))); return; }
                if (mode != "similar") return;
                var eids = Enumerable.Range(0, form.ResultCount).Select(form.ResultEid).ToList();
                if (step == 2)
                {
                    bool ok = eids.Count > 0 && eids[0] == expected && !eids.Contains(eid) && form.ResultConfidence(0) == "yüksek" &&
                              form.ResultWhy(0)[0].StartsWith("benzerlik ") && !form.ModelLoaded && form.StatusText.Contains("benzer e-posta");
                    Console.WriteLine("similar results={0} first={1} conf={2} model_loaded={3} ms={4} -> {5}", eids.Count, eids.FirstOrDefault(),
                                      eids.Count > 0 ? form.ResultConfidence(0) : "-", form.ModelLoaded, sw.ElapsedMilliseconds, ok ? "ok" : "FAIL");
                    if (!ok) { form.Shutdown(); return; }
                    // Right-click "Yalnız bu gönderenle yeniden ara" on the first result: "Benzerlerini bul" of eid again
                    // (not a text search), with the Gönderen filter of that result: every result has its sender.
                    form.SelectResult(0);
                    sender = src.Read(eids[0], storeId, "")["sender"] as string;
                    if (!form.FilterBySenderEnabled) { Console.WriteLine("FAIL: menu item off after a finished similar search"); form.Shutdown(); return; }
                    step = 3; sw.Restart();
                    form.BeginInvoke((Action)(() =>
                    {
                        form.ClickFilterBySender();
                        // A similar search runs now: the menu item is off (a click would not start anything).
                        bool off = form.SessionMode == "similar" && !form.FilterBySenderEnabled;
                        Console.WriteLine("filter_by_sender_restart mode={0} item_off_while_running={1} -> {2}", form.SessionMode, off, off ? "ok" : "FAIL");
                        if (!off) { form.Shutdown(); }
                    }));
                    return;
                }
                bool fok = eids.Count > 0 && !eids.Contains(eid) && eids.All(e => (src.Read(e, storeId, "")["sender"] as string) == sender) &&
                           form.StatusText.Contains("Filtre: gönderen") && form.FilterBoxes[0].Text.StartsWith(sender + " <");
                Console.WriteLine("similar_filtered results={0} same_sender={1} -> {2}", eids.Count, fok, fok ? "ok" : "FAIL");
                code = fok ? 0 : 1;
                Console.WriteLine(code == 0 ? "PASS" : "FAIL");
                form.Shutdown();
            };
            Application.Run(form);
            return code;
        }

        /// <summary>True when every string in the timing log row is under the key "at" or "reason" (no text of a mail or query).</summary>
        static bool NumbersOnly(object v, string key)
        {
            var d = v as Dictionary<string, object>;
            if (d != null) return d.All(kv => NumbersOnly(kv.Value, kv.Key));
            var a = v as ArrayList;
            if (a != null) return a.Cast<object>().All(x => NumbersOnly(x, key));
            return !(v is string) || key == "at" || key == "reason";
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
            // A started search is the first recent search, also in the file (the text is compared, never printed).
            var h = SearchHistory.Load(settings.Data);
            bool recorded = mode != "search" || h.Recent.Count > 0 && h.Recent[0] == form.QueryText.Trim() && form.History.Recent.FirstOrDefault() == h.Recent[0];
            Console.WriteLine("history_recorded=" + recorded);
            bool ok = !attLeft && pids.Count > 0 && indexFound && recorded;
            if (mode == "search")
            {
                // Timing log: the search wrote one JSON line with numbers only (the strings: "at" and "reason").
                string tl = Path.Combine(settings.Data, "timing.log");
                string last = File.Exists(tl) ? File.ReadAllLines(tl).LastOrDefault() : null;
                Dictionary<string, object> row = null;
                try { if (last != null) row = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<Dictionary<string, object>>(last); } catch { }
                bool tok = row != null && row.ContainsKey("manifest") && row.ContainsKey("worker_timing") && (string)row["reason"] == reason &&
                           (reason != "ok" || row.ContainsKey("final_ms")) && NumbersOnly(row, null);
                Console.WriteLine("timing_log lines={0} manifest_folders={1} final_ms={2} -> {3}", File.Exists(tl) ? File.ReadAllLines(tl).Length : 0,
                                  row != null && row.ContainsKey("manifest") ? ((ArrayList)row["manifest"]).Count : -1,
                                  row != null && row.ContainsKey("final_ms") ? row["final_ms"] : "-", tok ? "ok" : "FAIL");
                ok &= tok;
            }
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

        /// <summary>A mail source with a slow Read (a slow Outlook or a large mailbox) and mails that are not there yet (Hidden:
        /// not in the folder lists, "--update"): every other call is passed on.</summary>
        class SlowSource : IMailSource
        {
            readonly IMailSource s; readonly int delay;
            public readonly HashSet<string> Hidden = new HashSet<string>();
            public SlowSource(IMailSource s, int delay) { this.s = s; this.delay = delay; }
            public List<KeyValuePair<string, string>> Stores() { return s.Stores(); }
            public List<FolderInfo> Folders(string storeId) { return s.Folders(storeId); }
            public IEnumerable<object[]> Manifest(FolderInfo f, Func<bool> cancelled) { return s.Manifest(f, cancelled).Where(r => !Hidden.Contains((string)r[0])); }
            public Dictionary<string, object> Read(string eid, string storeId, string folderId)
            {
                if (delay > 0) System.Threading.Thread.Sleep(delay);
                return s.Read(eid, storeId, folderId);
            }
            public bool SaveAttachment(string eid, string storeId, int idx, string path) { return s.SaveAttachment(eid, storeId, idx, path); }
            public string AttachmentName(string eid, string storeId, int idx, out bool item) { return s.AttachmentName(eid, storeId, idx, out item); }
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
            if (File.Exists(SearchHistory.PathIn(settings.Data))) File.Delete(SearchHistory.PathIn(settings.Data)); // the window opens with no searches
            var form = new SearchForm(src, settings) { StartPosition = FormStartPosition.Manual, Location = new Point(-32000, -32000), ShowInTaskbar = false };
            int code = 1;
            form.Shown += (s, e) =>
            {
                // The window opens with the focus in the search box (before any check moves it).
                bool focusOpen = form.Active == form.QueryBox;
                Console.WriteLine("key_focus_open -> " + (focusOpen ? "ok" : "FAIL"));
                bool ok = focusOpen;
                // Fake worker events with realistic sizes (no worker in this check).
                form.Feed("index", Coverage(form));
                form.Feed("index", new Dictionary<string, object> { {"ev", "progress"}, {"phase", "embed"}, {"done", 1234}, {"total", 4321}, {"rate", 0.6}, {"eta_s", 5148} });
                form.Feed("search", Results());
                form.Feed("status", Facets(false)); // filter lists with long names (layout)
                // All filter rows: "Özel aralık…" shows the third row with the two date fields (cleared after the screenshots).
                form.FilterBoxes[3].SelectedIndex = form.FilterBoxes[3].Items.Count - 1; form.FilterBoxes[4].SelectedIndex = 3;
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
                        ok &= A11yCheck(root, name);
                        if (tab == 0) ok &= FilterGrid(form, name);
                        if (shotDir != null) Shot(root, Path.Combine(shotDir, "ui-" + name + ".png"));
                    }
                }
                form.ClickClearFilters();
                form.ShowTab(2);
                ok &= SettingsRoundTrip(form, settings);
                form.ShowTab(0);
                // "Eki aç" bar: only for a match inside an attachment; the button needs the attachment index.
                string l1 = form.AttachLine; bool b1 = form.AttachButtonEnabled;
                form.SelectResult(1); Application.DoEvents();
                string l2 = form.AttachLine;
                form.SelectResult(2); Application.DoEvents();
                string l3 = form.AttachLine; bool b3 = form.AttachButtonEnabled;
                form.SelectResult(0); Application.DoEvents();
                bool bar = l1 != null && l1.StartsWith("Ekte bulundu: Kargo_Faturalari_2026.zip › Eylül/") && b1 && l2 == null &&
                           l3 == "Ekte bulundu: Sayim_Eylul.xlsx" && !b3 && form.AttachLine == l1;
                Console.WriteLine("attach_bar zip_member={0} body_hidden={1} old_result_button_off={2} -> {3}", l1 != null && b1, l2 == null, l3 != null && !b3, bar ? "ok" : "FAIL");
                ok &= bar;
                Console.WriteLine("client_sizes default={0} min={1} root={2}", form.DesignClientSize, form.MinClientSize, root.Size);
                Console.WriteLine("preview_marked_chars=" + form.PreviewMarkedChars);
                ok &= form.PreviewMarkedChars > 0;
                ok &= FilterUiCheck(form);
                ok &= SortUiCheck(form);
                ok &= StageUiCheck(form);
                ok &= SearchListCheck(form, shotDir);
                ok &= HistoryUiCheck(form, settings, shotDir);
                ok &= IndexDeleteCheck(form, settings); // it clears the results
                ok &= ExcludeUiCheck(form, settings, src.Stores()[0].Key, shotDir);
                // Accessibility (the 7 features of 10-26): MSAA names, rows, high contrast, focus frame, keys, text size, DPI.
                form.Feed("search", Results());
                ok &= MsaaCheck(form);
                ok &= ColorRowRingCheck(form, shotDir);
                ok &= KeyCheck(form, (JsonSource)src, shotDir);
                ok &= ScaleCheck(form, settings, shotDir);
                ok &= DpiCheck(form, shotDir);
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
            var att = all.OfType<CheckBox>().First(x => x.AccessibleName == "Ek dosyalarının içinde de ara");
            var save = all.OfType<Button>().First(x => x.Text == "&Kaydet");
            bool ok = true;
            var first = all.OfType<CheckBox>().First(x => x.AccessibleName == "Arama bitince ilk sonuca geç");
            var battery = all.OfType<CheckBox>().First(x => x.AccessibleName == "Pille çalışırken indekslemeyi duraklat");
            foreach (var v in new[] { new { T = 6, M = 6144, B = 45, A = false, K = 30, C = 7, P = false, F = true }, new { T = 2, M = 3072, B = 0, A = true, K = 0, C = 0, P = true, F = false } })
            {
                num("İşlemci çekirdeği").Value = v.T; num("Bellek sınırı (MB)").Value = v.M; num("Aramada indeksleme (sn)").Value = v.B;
                num("Yapay zekâyı hazır tut (dk)").Value = v.K; num("Pencere kapandıktan sonra hazır tut (dk)").Value = v.C;
                att.Checked = v.A; battery.Checked = v.P; first.Checked = v.F;
                save.PerformClick();
                var r = Settings.Load();
                bool same = r.Threads == v.T && r.MemoryLimitMb == v.M && r.SearchIndexBudgetS == v.B && r.Attachments == v.A && r.KeepAliveMin == v.K &&
                            r.KeepAfterCloseMin == v.C && r.PauseOnBattery == v.P && r.FocusFirstResult == v.F && r.TextScalePct == 100 &&
                            form.StatusText.StartsWith("Ayarlar kaydedildi");
                Console.WriteLine("settings_save threads={0} memory={1} budget={2} attachments={3} keep_alive_min={4} after_close_min={5} battery={6} first_result={7} text={8} -> {9}",
                                  r.Threads, r.MemoryLimitMb, r.SearchIndexBudgetS, r.Attachments, r.KeepAliveMin, r.KeepAfterCloseMin, r.PauseOnBattery,
                                  r.FocusFirstResult, r.TextScalePct, same ? "ok" : "FAIL");
                ok &= same;
            }
            File.WriteAllText(settings.ConfigFile, "{\"threads\": 99, \"memory_limit_mb\": 10, \"search_index_budget_s\": -5, \"embed_budget_s\": 7, \"keep_alive_min\": 999, \"text_scale_pct\": 999}");
            var c = Settings.Load();
            bool clamp = c.Threads == 16 && c.MemoryLimitMb == 2048 && c.SearchIndexBudgetS == 0 && c.KeepAliveMin == 120 && c.TextScalePct == 200 &&
                         Settings.ClampTextScale(110) == 100 && Settings.ClampTextScale(130) == 125 && Settings.ClampTextScale(150) == 150 &&
                         Settings.ClampTextScale(-5) == 100 && Settings.ClampTextScale(188) == 200;
            Console.WriteLine("settings_clamp threads={0} memory={1} budget={2} keep_alive_min={3} text={4} -> {5}", c.Threads, c.MemoryLimitMb, c.SearchIndexBudgetS,
                              c.KeepAliveMin, c.TextScalePct, clamp ? "ok" : "FAIL");
            // A file without the new key (an older install): keep-alive is on with the default.
            File.WriteAllText(settings.ConfigFile, "{\"threads\": 4}");
            var old = Settings.Load();
            bool dflt = old.KeepAliveMin == 10 && old.TextScalePct == 100 && !old.FocusFirstResult;
            Console.WriteLine("settings_keep_alive_default=" + (dflt ? "ok" : "FAIL"));
            clamp &= dflt;
            bool tmpLeft = File.Exists(settings.ConfigFile + ".tmp");
            Console.WriteLine("settings_tmp_left=" + tmpLeft);
            File.Delete(settings.ConfigFile);
            return ok && clamp && !tmpLeft;
        }

        // Recent and saved searches, file rules only (no window: runs on mono on the Mac too). Works in a new folder that
        // it makes under the test home and deletes at the end. Prints flags only.
        static bool HistoryCheck(Settings settings)
        {
            string dir = Path.Combine(settings.Home, "history-check-" + Guid.NewGuid().ToString("N"));
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("history_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            try
            {
                var h = SearchHistory.Load(dir); // the folder is not there yet
                check("empty", h.Recent.Count == 0 && h.Saved.Count == 0 && !h.ReadFailed && !Directory.Exists(dir));
                h.AddRecent("Kargo faturası"); h.AddRecent("   "); h.AddRecent(null); h.AddRecent(" KARGO faturası ");
                check("dedupe", h.Recent.Count == 1 && h.Recent[0] == "KARGO faturası");
                // Turkish ı/I, i/İ: one or two lines, by platform (see SearchHistory.AddRecent). Only "no crash, no loss".
                h.AddRecent("kargo FATURASI");
                check("dedupe_turkish", h.Recent.Count >= 1 && h.Recent[0] == "kargo FATURASI");
                for (int i = 0; i < 25; i++) h.AddRecent("q" + i);
                check("limit", h.Recent.Count == SearchHistory.MaxRecent && h.Recent[0] == "q24" && h.Recent[19] == "q5");
                string odd = "Ahmet'in \"kargo\" & fatura – İzmir\\Şube";
                h.AddRecent(odd);
                bool added = h.SetSaved("Fatura", "kargo faturası"), updated = h.SetSaved(" fatura ", odd);
                bool noEmpty = !h.SetSaved("", "a") && !h.SetSaved("b", "  ") && !h.SetSaved(null, null);
                check("saved", !added && updated && noEmpty && h.Saved.Count == 1 && h.Saved[0].Key == "fatura" && h.Saved[0].Value == odd);
                var c = h.Clone(); c.Recent.Clear(); c.Saved.Clear();
                check("clone", h.Recent.Count == 20 && h.Saved.Count == 1);
                h.Save(); // makes the folder
                var r = SearchHistory.Load(dir);
                check("roundtrip", r.Recent.SequenceEqual(h.Recent) && r.Saved.SequenceEqual(h.Saved) && !File.Exists(h.FilePath + ".tmp"));
                File.WriteAllText(h.FilePath, "{not json");
                r = SearchHistory.Load(dir);
                check("corrupt", r.Recent.Count == 0 && r.Saved.Count == 0 && !r.ReadFailed);
                File.WriteAllText(h.FilePath, "[1, 2]");
                r = SearchHistory.Load(dir);
                check("not_object", r.Recent.Count == 0 && r.Saved.Count == 0);
                File.WriteAllText(h.FilePath, "{\"recent\": [1, \"a\", null, \"A\"], \"saved\": [{\"name\": 5, \"query\": \"q\"}, {\"name\": \"n\", \"query\": \"q\"}, \"x\", {\"name\": \"m\"}]}");
                r = SearchHistory.Load(dir);
                check("wrong_types", r.Recent.Count == 1 && r.Recent[0] == "a" && r.Saved.Count == 1 && r.Saved[0].Key == "n");
                File.WriteAllText(h.FilePath, "{\"recent\": \"a\", \"saved\": {}}");
                r = SearchHistory.Load(dir);
                check("wrong_containers", r.Recent.Count == 0 && r.Saved.Count == 0);
                // A write that fails (the file name is a folder) throws and leaves no temp copy of the search texts.
                File.Delete(h.FilePath); Directory.CreateDirectory(h.FilePath);
                bool threw = false;
                try { h.Save(); } catch (IOException) { threw = true; } catch (UnauthorizedAccessException) { threw = true; }
                check("write_error", threw && !File.Exists(h.FilePath + ".tmp"));
            }
            catch (Exception ex) { Console.WriteLine("history_check exception " + ex.GetType().Name + ": " + ex.Message); ok = false; }
            finally { if (Directory.Exists(dir)) Directory.Delete(dir, true); } // only the folder made above
            Console.WriteLine(ok ? "PASS" : "FAIL");
            return ok;
        }

        // "Eki aç" file rules (no window: runs on mono on the Mac too): safe file names, the zip member extraction with the
        // name rules of Python's zipfile (the worker names the member), the 20 MB member cap, read-only copy and cleanup.
        // Works in a new folder under the test home and deletes it at the end. Prints flags and file names only.
        static bool AttachCheck(string corpus, Settings settings)
        {
            string dir = Path.Combine(settings.Home, "attach-check-" + Guid.NewGuid().ToString("N"));
            string fixture = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(corpus)), "att", "Bakim_Arsivi.zip");
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("attach_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            Func<string, string> copy = sub => { string d = Path.Combine(dir, sub); Directory.CreateDirectory(d); string p = Path.Combine(d, "Bakim_Arsivi.zip"); File.Copy(fixture, p); return p; };
            try
            {
                Func<string, string> safe = AttachmentCopy.SafeFileName;
                string longName = new string('a', 200) + ".pdf";
                check("safe_name", safe("a/b\\c:d*?.txt") == "a_b_c_d__.txt" && safe("CON.txt") == "_CON.txt" && safe("lpt1") == "_lpt1" &&
                                   safe("  ... ") == "ek" && safe(null) == "ek" && safe("rapor. ") == "rapor" && safe("Fatura Ekim.pdf") == "Fatura Ekim.pdf" &&
                                   safe(longName).Length == 120 && safe(longName).EndsWith(".pdf") &&
                                   safe(new string('x', 150) + "." + new string('b', 30)) == new string('x', 120)); // a long "extension" is cut too
                var src = new JsonSource(corpus);
                bool item;
                check("source_name", src.AttachmentName("A13", "", 1, out item) == "Bakim_Arsivi.zip" && !item && src.AttachmentName("A13", "", 2, out item) == null);
                // The member of the eval fixture: the bytes of the docx inside the zip; the zip copy is deleted.
                string zip = copy("a");
                string m = AttachmentCopy.ExtractMember(zip, "Bakim/bakim_notu.docx");
                check("member", m != null && Path.GetFileName(m) == "bakim_notu.docx" && Path.GetDirectoryName(m) == Path.GetDirectoryName(zip) &&
                                Sha256(m) == "0136fc85186c56d6352db34b79f830630b09fcf086d9150c0883491002ce610d" && !File.Exists(zip));
                zip = copy("b");
                check("member_missing", AttachmentCopy.ExtractMember(zip, "Bakim/yok.txt") == null && AttachmentCopy.ExtractMember(zip, "Bakim/") == null &&
                                        File.Exists(zip) && Directory.GetFiles(Path.GetDirectoryName(zip)).Length == 1);
                // Names as other zip programs write them. Flag off (Windows Explorer: OEM code page, 857 on a Turkish PC):
                // Python reads the bytes as cp437. Flag on: UTF-8. A "\" in a name: the worker on Windows has "/".
                Encoding oem;
                try { oem = Encoding.GetEncoding(857); } catch { oem = Encoding.GetEncoding(437); }
                // Two files: an Update would write every name again with the UTF-8 flag.
                string tr = "Ölçüm İrsaliye ğş.txt", zOem = Path.Combine(dir, "c", "oem.zip"), zUtf = Path.Combine(dir, "c", "utf8.zip");
                Directory.CreateDirectory(Path.GetDirectoryName(zOem));
                using (var z = new ZipArchive(File.Create(zOem), ZipArchiveMode.Create, false, oem))
                {
                    Write(z.CreateEntry("Klasör/" + tr), "oem");
                    Write(z.CreateEntry("Dir\\b.txt"), "backslash");
                }
                using (var z = new ZipArchive(File.Create(zUtf), ZipArchiveMode.Create)) Write(z.CreateEntry("UTF/" + tr), "utf8"); // default: UTF-8 flag
                string py = Encoding.GetEncoding(437).GetString(oem.GetBytes("Klasör/" + tr));
                foreach (var t in new[] { new[] { py, "oem", zOem }, new[] { "UTF/" + tr, "utf8", zUtf }, new[] { "Dir/b.txt", "backslash", zOem } })
                {
                    string c = Path.Combine(dir, "c", t[1]); Directory.CreateDirectory(c);
                    string zc = Path.Combine(c, "x.zip"); File.Copy(t[2], zc);
                    string f = AttachmentCopy.ExtractMember(zc, t[0]);
                    check("name_" + t[1], f != null && File.ReadAllText(f) == t[1] && Path.GetExtension(f) == ".txt");
                    if (f != null) Console.WriteLine("attach_name_" + t[1] + "_file=" + Path.GetFileName(f));
                }
                // A member bigger than the cap (a zip bomb, or a false size in the header): stop, no part file stays.
                string z3 = Path.Combine(dir, "d", "big.zip");
                Directory.CreateDirectory(Path.GetDirectoryName(z3));
                using (var z = new ZipArchive(File.Create(z3), ZipArchiveMode.Create))
                using (var s = z.CreateEntry("big.txt", CompressionLevel.Optimal).Open())
                {
                    var buf = new byte[1 << 20];
                    for (int i = 0; i < 20; i++) s.Write(buf, 0, buf.Length);
                    s.WriteByte(0);
                }
                bool threw = false;
                try { AttachmentCopy.ExtractMember(z3, "big.txt"); } catch (InvalidDataException) { threw = true; }
                check("member_cap", threw && !File.Exists(Path.Combine(dir, "d", "big.txt")) && File.Exists(z3));
                // The opened copy is read-only; Clean deletes it (and a read-only file) and the folder.
                string root = AttachmentCopy.Root(dir), one = Path.Combine(root, "1234abcd");
                Directory.CreateDirectory(one);
                string p = Path.Combine(one, "x.txt");
                File.WriteAllText(p, "x");
                AttachmentCopy.Protect(p);
                bool ro = (File.GetAttributes(p) & FileAttributes.ReadOnly) != 0;
                string zone = ZoneOf(p);
                bool motw = Environment.OSVersion.Platform != PlatformID.Win32NT || zone != null && zone.Contains("ZoneId=3");
                AttachmentCopy.Clean(dir);
                check("protect_clean", ro && motw && !Directory.Exists(root));
            }
            catch (Exception ex) { Console.WriteLine("attach_check exception " + ex.GetType().Name + ": " + ex.Message); ok = false; }
            finally
            {
                if (Directory.Exists(dir))
                {
                    foreach (var f in Directory.GetFiles(dir, "*", SearchOption.AllDirectories)) File.SetAttributes(f, FileAttributes.Normal);
                    Directory.Delete(dir, true); // only the folder made above
                }
            }
            Console.WriteLine(ok ? "PASS" : "FAIL");
            return ok;
        }

        static void Write(ZipArchiveEntry e, string text) { using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false))) w.Write(text); }

        static string Sha256(string path)
        {
            using (var sha = SHA256.Create()) using (var f = File.OpenRead(path))
                return BitConverter.ToString(sha.ComputeHash(f)).Replace("-", "").ToLowerInvariant();
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateFile(string name, uint access, uint share, IntPtr security,
                                                                            uint disposition, uint flags, IntPtr template);

        /// <summary>The Zone.Identifier stream of a file (Mark of the Web), or null (none, or not Windows).</summary>
        static string ZoneOf(string path)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return null;
            using (var h = CreateFile(path + ":Zone.Identifier", 0x80000000, 1, IntPtr.Zero, 3, 0, IntPtr.Zero)) // GENERIC_READ, OPEN_EXISTING
            {
                if (h.IsInvalid) return null;
                using (var r = new StreamReader(new FileStream(h, FileAccess.Read), Encoding.ASCII)) return r.ReadToEnd();
            }
        }

        // Ara tab: recent and saved searches in the window (Windows only: WinForms). Starts with no searches file (made by
        // LayoutCheck) and deletes the file at the end. No worker in the layout check: Ara stops at "Arama bileşeni
        // bulunamadı", so a search from the menu fills the box but is not recorded (only started searches are).
        static bool HistoryUiCheck(SearchForm form, Settings settings, string shotDir)
        {
            string file = SearchHistory.PathIn(settings.Data);
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("history_ui_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            Func<List<ToolStripItem>, string, ToolStripMenuItem> item = (items, text) => items.OfType<ToolStripMenuItem>().FirstOrDefault(x => x.Text == text);
            check("start_empty", form.History.Recent.Count == 0 && form.History.Saved.Count == 0 && form.Suggestions.Length == 0);
            string longQ = "Ahmet & Ayşe – sözleşme taslağı fesih bildirim süresi ve ekleri hakkında uzun bir sorgu";
            form.RememberSearch("kargo faturası"); form.RememberSearch(longQ);
            form.SetQuery("kargo faturası eylül");
            bool saved = form.SaveSearchAs("Aylık kargo") && form.StatusText.StartsWith("Arama kaydedildi");
            form.SetQuery("kargo faturası ekim");
            bool upd = form.SaveSearchAs("aylık KARGO") && form.StatusText.StartsWith("Kayıtlı arama güncellendi");
            Application.DoEvents(); // the suggestion list changes after the current message
            var sug = form.Suggestions;
            var disk = SearchHistory.Load(settings.Data);
            check("save", saved && upd && form.History.Saved.Count == 1 && form.History.Saved[0].Value == "kargo faturası ekim" &&
                          disk.Saved.SequenceEqual(form.History.Saved) && disk.Recent.SequenceEqual(form.History.Recent) && form.History.Recent.Count == 2);
            check("suggestions", sug.Length == 3 && sug.Contains("kargo faturası") && sug.Contains(longQ) && sug.Contains("kargo faturası ekim"));
            var menu = form.HistoryMenuItems();
            var run = item(menu, "aylık KARGO");
            var recentLong = menu.OfType<ToolStripMenuItem>().FirstOrDefault(x => x.ToolTipText == longQ);
            check("menu", run != null && run.Enabled && run.ToolTipText == "kargo faturası ekim" && item(menu, "Bu aramayı &kaydet…").Enabled &&
                          item(menu, "Kayıtlı aramayı &sil").DropDownItems.Count == 1 && item(menu, "Son aramaları &temizle").Enabled &&
                          recentLong != null && recentLong.Text == SearchForm.MenuText(longQ) && recentLong.Text.Contains("&&") && recentLong.Text.EndsWith("…") &&
                          recentLong.Text.Replace("&&", "&").Length == 60);
            form.SetQuery("");
            run.PerformClick(); // one click: the query goes into the box and Ara runs (it stops here: no worker installed)
            check("run_click", form.QueryText == "kargo faturası ekim" && form.History.Recent.Count == 2 && form.StatusText.StartsWith("Arama bileşeni bulunamadı"));
            // Name dialog: layout at 2x, Kaydet off for an empty name.
            using (var d = new SearchForm.NameDialog(form, "Aylık kargo") { StartPosition = FormStartPosition.Manual, Location = new Point(-32000, -32000) })
            {
                d.Show(); Application.DoEvents(); // no owner: Show(owner) needs an interactive session (the SSH session on 2.19 is not)
                bool fits = Check(d, "name-dialog") & A11yCheck(d, "name-dialog");
                if (shotDir != null) Shot(d.Body, Path.Combine(shotDir, "ui-name-dialog.png"));
                var okBtn = SearchForm.AllControls(d).OfType<Button>().First(x => x.DialogResult == DialogResult.OK);
                d.Box.Text = "  "; bool off = !okBtn.Enabled;
                d.Box.Text = "x"; bool on = okBtn.Enabled;
                check("name_dialog", fits && off && on && d.Box.MaxLength == 80 && d.AcceptButton == okBtn);
                d.Close();
            }
            menu = form.HistoryMenuItems();
            item(menu, "Son aramaları &temizle").PerformClick();
            disk = SearchHistory.Load(settings.Data);
            check("clear_recent", form.History.Recent.Count == 0 && form.History.Saved.Count == 1 && disk.Recent.Count == 0 && disk.Saved.Count == 1 &&
                                  form.StatusText.StartsWith("Son aramalar silindi"));
            menu = form.HistoryMenuItems();
            item(menu, "Kayıtlı aramayı &sil").DropDownItems[0].PerformClick();
            menu = form.HistoryMenuItems();
            check("delete_saved", form.History.Saved.Count == 0 && SearchHistory.Load(settings.Data).Saved.Count == 0 && item(menu, "(kayıtlı arama yok)") != null &&
                                  !item(menu, "Kayıtlı aramayı &sil").Enabled && !item(menu, "Son aramaları &temizle").Enabled);
            // A file locked by another program at the window open: no write over it (that would delete its searches).
            form.RememberSearch("önceki arama");
            bool blocked;
            using (new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                form.ReloadHistory();
                blocked = form.History.ReadFailed && !form.SaveSearchAs("y") && form.StatusText.Contains("aramalar dosyası okunamadı");
            }
            bool kept = form.SaveSearchAs("y") && form.History.Recent.SequenceEqual(new[] { "önceki arama" }) && SearchHistory.Load(settings.Data).Saved.Count == 1;
            check("locked_file", blocked && kept);
            // A write error: the window keeps the saved state, no temp copy stays.
            File.Delete(file); Directory.CreateDirectory(file);
            form.SetQuery("z");
            bool failed = !form.SaveSearchAs("z") && form.StatusText.StartsWith("Arama kaydedilemedi") && form.History.Saved.Count == 1 && !File.Exists(file + ".tmp");
            Directory.Delete(file);
            check("write_error", failed);
            if (File.Exists(file)) File.Delete(file);
            form.ReloadHistory(); Application.DoEvents();
            return ok;
        }

        // Ara filters (Windows only: WinForms): the lists from done.facets in Turkish order, the chosen value kept by its key
        // when the lists come again, "(tümü)" when it is gone; the folder list = the checked İndeks folders; FilteredScope
        // gives the folder scope, the keys for the worker and the status note. The start state comes back.
        static bool FilterUiCheck(SearchForm form)
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("filter_ui_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            var b = form.FilterBoxes;
            Func<ComboBox, string[]> texts = x => x.Items.Cast<object>().Select(o => o.ToString()).ToArray();
            check("lists", texts(b[0]).SequenceEqual(new[] { "(tümü)", "Ahmet Demir <ahmet.demir@hizli-lojistik.example>", "Çağla Öztürk <cagla@x.example>", "Zeynep Arslan <zeynep@x.example>" }) &&
                           texts(b[1]).SequenceEqual(new[] { "(tümü)", "Furkan Çakır", "Muhasebe Ekibi" }) && b.All(x => x.SelectedIndex == 0) &&
                           texts(b[2]).Length == form.ListedFolderIds().Count + 1);
            b[0].SelectedIndex = 3; b[1].SelectedIndex = 1;
            bool changed = form.StatusText.StartsWith("Filtre değişti");
            var fb = (SearchForm.SearchList)b[2];
            fb.SetChecked(new[] { "F-PROJ" });
            bool one = fb.Text == "Gelen Kutusu\\Projeler" && form.ClearFiltersEnabled;
            form.Feed("status", Facets(true)); // the lists come again: Zeynep stays, Furkan Çakır is gone
            Dictionary<string, object> f; string note;
            var scope = form.FilteredScope(form.AllFolders, out f, out note);
            check("kept", changed && one && b[0].Text == "Zeynep Arslan <zeynep@x.example>" && b[1].SelectedIndex == 0 &&
                          (string)f["sender"] == "zeynep@x.example" && f["recipient"] == null && scope.Count == 1 && scope[0].Id == "F-PROJ" &&
                          note == " Filtre: gönderen Zeynep Arslan <zeynep@x.example>, klasör Gelen Kutusu\\Projeler.");
            // More than one folder: the scope is these folders; the box shows the count, the status note the names.
            fb.SetChecked(new[] { "F-PROJ", "F-ARCH" });
            scope = form.FilteredScope(form.AllFolders, out f, out note);
            check("folders_multi", fb.Text == "2 klasör seçili" && scope.Select(x => x.Id).OrderBy(x => x).SequenceEqual(new[] { "F-ARCH", "F-PROJ" }) &&
                                   note.Contains("klasör Arşiv, Gelen Kutusu\\Projeler."));
            form.SetFolderChecked("F-PROJ", false); // not in the scope any more: that folder leaves the filter, Arşiv stays
            bool gone = !texts(b[2]).Contains("Gelen Kutusu\\Projeler") && fb.Checked.SetEquals(new[] { "F-ARCH" }) && fb.Text == "Arşiv";
            form.SetFolderChecked("F-PROJ", true);
            check("folder_unchecked", gone && texts(b[2]).Contains("Gelen Kutusu\\Projeler") && fb.Checked.SetEquals(new[] { "F-ARCH" }));
            foreach (var x in b) x.SelectedIndex = 0;
            fb.SetChecked(new string[0]);
            form.FilteredScope(form.AllFolders, out f, out note);
            check("none", note == "" && f["sender"] == null && f["recipient"] == null && f["attach"] == null && !f.ContainsKey("date_from") &&
                          !form.ClearFiltersEnabled);
            // Tarih: a preset gives [from, to) in local time; "Özel aralık…" shows the two pickers, includes both days and
            // turns a reversed range. Ek: the worker key. "Filtreleri temizle" sets every list back to "(tümü)".
            Func<ComboBox, string, int> pick = (x, text) => texts(x).ToList().IndexOf(text);
            int y = DateTime.Today.Year;
            b[3].SelectedIndex = pick(b[3], "Geçen yıl"); b[4].SelectedIndex = pick(b[4], "PDF");
            form.FilteredScope(form.AllFolders, out f, out note);
            check("date_preset", (string)f["date_from"] == y - 1 + "-01-01T00:00:00" && (string)f["date_to"] == y + "-01-01T00:00:00" &&
                                 (string)f["attach"] == "pdf" && note == " Filtre: tarih Geçen yıl, ek PDF." && form.ClearFiltersEnabled &&
                                 !form.DatePickers[0].Visible);
            b[3].SelectedIndex = pick(b[3], "Özel aralık…");
            form.DatePickers[0].Value = new DateTime(2026, 9, 30); form.DatePickers[1].Value = new DateTime(2026, 9, 1);
            form.FilteredScope(form.AllFolders, out f, out note);
            check("date_custom", form.DatePickers.All(p => p.Visible) && (string)f["date_from"] == "2026-09-01T00:00:00" &&
                                 (string)f["date_to"] == "2026-10-01T00:00:00" && note.Contains("tarih 01.09.2026–30.09.2026"));
            b[0].SelectedIndex = 1; fb.SetChecked(new[] { "F-ARCH" });
            form.ClickClearFilters();
            form.FilteredScope(form.AllFolders, out f, out note);
            check("clear", b.All(x => x.SelectedIndex == 0) && fb.Checked.Count == 0 && fb.Text == "(tümü)" && note == "" && !form.ClearFiltersEnabled &&
                           !form.DatePickers[0].Visible);
            return ok;
        }

        // Filter fields (Windows only): the same width in each column, the same left edge in each column, the same top and
        // height in each row. The date fields of the third row line up with the lists above them.
        static bool FilterGrid(SearchForm form, string name)
        {
            var b = form.FilterBoxes;
            var d = form.DatePickers;
            var cols = new[] { new Control[] { b[0], b[3], d[0] }, new Control[] { b[1], b[4], d[1] } };
            var rows = new[] { new Control[] { b[0], b[1], b[2] }, new Control[] { b[3], b[4] }, new Control[] { d[0], d[1] } };
            Func<Control, Rectangle> at = c => form.Root.RectangleToClient(c.RectangleToScreen(new Rectangle(Point.Empty, c.Size)));
            bool ok = cols.All(c => c.Where(x => x.Visible).Select(x => at(x).Left).Distinct().Count() == 1 && c.Where(x => x.Visible).Select(x => x.Width).Distinct().Count() == 1) &&
                      new[] { b[0], b[1], b[2] }.Select(x => x.Width).Max() - new[] { b[0], b[1], b[2] }.Select(x => x.Width).Min() <= 2 &&
                      rows.All(r => r.Where(x => x.Visible).Select(x => at(x).Top).Distinct().Count() <= 1) &&
                      b.All(x => x.Height == b[0].Height) && d.All(x => x.Visible);
            // With all filter rows, the result list still shows its 3 results (the preview gives way in a small window).
            var lv = SearchForm.AllControls(form.Root).OfType<ListView>().First(x => x.AccessibleName == "Sonuçlar");
            bool rows3 = lv.Items.Count >= 3 && lv.GetItemRect(2).Bottom <= lv.ClientSize.Height;
            ok &= rows3;
            if (!rows3 && lv.Items.Count >= 3)
                Console.WriteLine("filter_grid {0} list_client_h={1} top_item={2} row0={3} row2={4} list_h={5}", name, lv.ClientSize.Height,
                                  lv.TopItem != null ? lv.TopItem.Index : -1, lv.GetItemRect(0), lv.GetItemRect(2), lv.Height);
            Console.WriteLine("filter_grid {0} widths={1} lefts={2} tops={3} list_shows_3_rows={4} -> {5}", name, string.Join("/", b.Concat<Control>(d).Select(x => x.Width)),
                              string.Join("/", b.Concat<Control>(d).Select(x => at(x).Left)), string.Join("/", b.Concat<Control>(d).Select(x => at(x).Top)), rows3, ok ? "ok" : "FAIL");
            return ok;
        }

        // Search field in the Gönderen list (Windows only): the drop-down opens with the field; words filter without case
        // and Turkish marks; arrows move; Enter takes the item (the Gönderen filter changes, the list closes); Esc closes
        // and keeps the value; a letter on the closed list opens it with that letter; no match gives a note that is not a
        // choice. The start state comes back.
        static bool SearchListCheck(SearchForm form, string shotDir)
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("search_list_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            form.Feed("status", Facets(false));
            var box = (SearchForm.SearchList)form.FilterBoxes[0];
            Func<string[]> shown = () => box.Matches.Items.Cast<object>().Select(o => o.ToString()).ToArray();
            box.ShowPopup("");
            Application.DoEvents();
            // No focus check: the SSH session on 2.19 is not interactive (a window there does not take the focus).
            check("open", box.PopupOpen && box.PopupVisible && shown().Length == 4 && shown()[0] == "(tümü)" && box.Matches.SelectedIndex == 0 &&
                          box.PopupBody.Width >= box.Width && box.PopupBody.Height >= box.Field.Height + 3 * box.Matches.ItemHeight);
            box.Field.Text = "cagla";
            check("fold", shown().SequenceEqual(new[] { "Çağla Öztürk <cagla@x.example>" }) && box.Matches.SelectedIndex == 0);
            box.Field.Text = "ÖZTÜRK x.example";
            check("words", shown().SequenceEqual(new[] { "Çağla Öztürk <cagla@x.example>" }));
            box.Field.Text = "example";
            box.FieldKey(Keys.Down);
            if (shotDir != null) Shot(box.PopupBody, Path.Combine(shotDir, "ui-search-list.png"));
            bool moved = box.Matches.SelectedItem?.ToString() == "Çağla Öztürk <cagla@x.example>";
            box.FieldKey(Keys.Enter);
            Application.DoEvents();
            check("enter", moved && !box.PopupOpen && !box.PopupVisible && box.Text == "Çağla Öztürk <cagla@x.example>" &&
                           form.StatusText.StartsWith("Filtre değişti") && form.ClearFiltersEnabled);
            box.ShowPopup("zzz");
            bool none = shown().SequenceEqual(new[] { "(eşleşen yok)" }) && box.Matches.SelectedIndex == -1;
            box.FieldKey(Keys.Enter);
            bool stays = box.PopupOpen && box.Text == "Çağla Öztürk <cagla@x.example>";
            box.FieldKey(Keys.Escape);
            Application.DoEvents();
            check("no_match_and_esc", none && stays && !box.PopupOpen && box.Text == "Çağla Öztürk <cagla@x.example>");
            // A letter on the closed list (the protected OnKeyPress, as the key message calls it).
            var kp = typeof(Control).GetMethod("OnKeyPress", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            kp.Invoke(box, new object[] { new KeyPressEventArgs('y') }); // only "Zeynep" has a "y"
            Application.DoEvents();
            check("type_to_open", box.PopupOpen && box.Field.Text == "y" && shown().SequenceEqual(new[] { "Zeynep Arslan <zeynep@x.example>" }) &&
                                  box.Text == "Çağla Öztürk <cagla@x.example>");
            box.FieldKey(Keys.Escape);
            form.ClickClearFilters();
            check("clear", box.SelectedIndex == 0 && !box.PopupOpen && !form.ClearFiltersEnabled);
            // Klasör: check boxes. A click or Enter turns a row on or off and the list stays open; "(tümü)" is on only
            // while no row is on; turning it on clears the rows. The filter applies at the next Ara (Filtre değişti).
            var fb = (SearchForm.SearchList)form.FilterBoxes[2];
            var clb = (CheckedListBox)fb.Matches;
            Func<string> on = () => string.Join("|", clb.CheckedItems.Cast<object>().Select(o => o.ToString()));
            fb.ShowPopup("");
            Application.DoEvents();
            bool start = fb.PopupOpen && on() == "(tümü)" && clb.Items.Count == form.ListedFolderIds().Count + 1;
            fb.Field.Text = "arsiv";
            fb.FieldKey(Keys.Enter);
            fb.Field.Text = "projeler";
            fb.FieldKey(Keys.Enter);
            fb.Field.Text = "";
            if (shotDir != null) Shot(fb.PopupBody, Path.Combine(shotDir, "ui-folder-list.png"));
            check("folders_check", start && fb.PopupOpen && on() == "Arşiv|Gelen Kutusu\\Projeler" && fb.Text == "2 klasör seçili" &&
                                   fb.Checked.SetEquals(new[] { "F-ARCH", "F-PROJ" }) && form.StatusText.StartsWith("Filtre değişti") && form.ClearFiltersEnabled);
            fb.Field.Text = "arsiv";
            fb.FieldKey(Keys.Enter); // off again
            fb.Field.Text = "";
            bool offOne = on() == "Gelen Kutusu\\Projeler" && fb.Text == "Gelen Kutusu\\Projeler";
            clb.SelectedIndex = 0;
            fb.FieldKey(Keys.Enter); // "(tümü)": every row off
            check("folders_all", offOne && on() == "(tümü)" && fb.Checked.Count == 0 && fb.Text == "(tümü)" && fb.PopupOpen);
            fb.FieldKey(Keys.Enter); // "(tümü)" alone does not turn off
            bool stay = on() == "(tümü)";
            fb.FieldKey(Keys.Escape);
            Application.DoEvents();
            check("folders_close", stay && !fb.PopupOpen && !form.ClearFiltersEnabled);
            return ok;
        }

        // Result list: a click on a column header sorts by it (▲/▼ in the header), a second click turns the order; Güven
        // = the worker's order; new results keep the chosen column. Right-click "Yalnız bu gönderenle yeniden ara": off
        // for a sender that is not in the Gönderen list, and off before the first search or "Benzerlerini bul" (no worker
        // in this check: nothing ran, so nothing to repeat and the filter stays). The menu run itself: SimilarRun step 3.
        static bool SortUiCheck(SearchForm form)
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("sort_ui_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            Func<string> order = () => string.Join(",", Enumerable.Range(0, form.ResultCount).Select(form.ResultEid));
            check("start", order() == "E1,E2,E3" && form.HeaderText(3) == "Güven ▼");
            form.ClickHeader(2);
            bool dateDesc = order() == "E3,E1,E2" && form.HeaderText(2) == "Tarih ▼" && form.HeaderText(3) == "Güven";
            form.ClickHeader(2);
            check("date", dateDesc && order() == "E2,E1,E3" && form.HeaderText(2) == "Tarih ▲");
            form.ClickHeader(0);
            check("subject", order() == "E3,E2,E1" && form.HeaderText(0) == "Konu ▲");
            form.Feed("search", Results());
            check("kept_for_new_results", order() == "E3,E2,E1");
            form.ClickHeader(3);
            check("relevance", order() == "E1,E2,E3" && form.HeaderText(3) == "Güven ▼");
            form.Feed("status", Facets(false)); // Ahmet Demir is in the Gönderen list again
            form.SelectResult(2);
            bool noChoice = !form.SelectedHasSenderChoice; // Ayşe Kaya: not in the list
            form.SelectResult(0);
            bool hasChoice = form.SelectedHasSenderChoice, enabled = form.FilterBySenderEnabled;
            form.ClickFilterBySender();
            check("filter_by_sender_off_before_a_run", noChoice && hasChoice && !enabled && form.FilterBoxes[0].SelectedIndex == 0);
            return ok;
        }

        // The early semantic list (stage "semantic") and the final list of one search: the status text tells that the list
        // comes again; the mail that the user selected stays selected in the next list; without a choice by the user the
        // first row of the new list is selected.
        static bool StageUiCheck(SearchForm form)
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("stage_ui_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            Func<string, string[], Dictionary<string, object>> res = (stage, order) =>
            {
                var r = Results();
                var byEid = ((ArrayList)r["results"]).Cast<Dictionary<string, object>>().ToDictionary(x => (string)x["eid"]);
                r["results"] = new ArrayList(order.Select(e => byEid[e]).ToList());
                r["stage"] = stage;
                return r;
            };
            if (form.HeaderText(3) != "Güven ▼") form.ClickHeader(3); // the worker order for this check
            // The keyword list of a new search: the selection of the check before (E1) is not in it.
            form.Feed("search", res("keyword", new[] { "E2", "E3" }));
            bool kw = form.SelectedEid == "E2" && form.StatusText.Contains("Anlamsal arama sürüyor");
            form.Feed("search", res("semantic", new[] { "E1", "E2", "E3" }));
            check("keyword_first_row", kw);
            check("semantic_status", form.StatusText.Contains("Bitince liste bir kez daha yenilenir") && form.SelectedEid == "E1");
            form.Feed("search", res("final", new[] { "E2", "E1", "E3" }));
            check("first_row_follows", form.SelectedEid == "E2" && form.StatusText.StartsWith("3 sonuç"));
            form.SelectResult(2); // the user selects E3
            form.Feed("search", res("final", new[] { "E2", "E3", "E1" }));
            check("user_choice_kept", form.SelectedEid == "E3");
            form.Feed("search", res("final", new[] { "E1", "E2" })); // E3 is not in the new list
            check("gone_choice_first_row", form.SelectedEid == "E1");
            form.Feed("search", Results());
            return ok;
        }

        static Dictionary<string, object> Facets(bool second)
        {
            Func<object[], ArrayList> row = a => new ArrayList(a);
            var senders = new ArrayList { row(new object[] { "zeynep@x.example", "Zeynep Arslan", "zeynep@x.example", 4 }),
                                          row(new object[] { "cagla@x.example", "Çağla Öztürk", "cagla@x.example", 2 }) };
            if (!second) senders.Insert(0, row(new object[] { "ahmet.demir@hizli-lojistik.example", "Ahmet Demir", "ahmet.demir@hizli-lojistik.example", 9 }));
            var recips = new ArrayList { row(new object[] { "muhasebe ekibi", "Muhasebe Ekibi", 3 }) };
            if (!second) recips.Add(row(new object[] { "furkan cakir", "Furkan Çakır", 12 }));
            return new Dictionary<string, object> { {"ev", "done"}, {"status", "ok"}, {"facets", new Dictionary<string, object> { {"senders", senders}, {"recipients", recips} }} };
        }

        // Ayarlar "İndeksi sil". The confirmation dialog is modal: the check calls DeleteIndex (the code behind "Evet").
        // A locked index file stops the delete without an exception; after the unlock, the second press deletes
        // data\index and data\tmp-att, and settings.json stays. Runs only when the test data folder has no index.
        static bool IndexDeleteCheck(SearchForm form, Settings settings)
        {
            string index = Path.Combine(settings.Data, "index"), att = Path.Combine(settings.Data, "tmp-att");
            if (Directory.Exists(index) || Directory.Exists(att)) { Console.WriteLine("index_delete: data folder has an index, not touched: " + settings.Data + " -> FAIL"); return false; }
            var btn = SearchForm.AllControls(form).OfType<Button>().FirstOrDefault(x => x.Text == "İndeksi &sil");
            Directory.CreateDirectory(index); Directory.CreateDirectory(att);
            string db = Path.Combine(index, "a.sqlite");
            foreach (var f in new[] { db, db + "-wal", db + "-shm", Path.Combine(att, "000001.bin") }) File.WriteAllText(f, "x");
            File.WriteAllText(settings.ConfigFile, "{\"threads\": 4}");
            string searches = SearchHistory.PathIn(settings.Data);
            File.WriteAllText(searches, "{}");
            bool locked;
            using (new FileStream(db, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                form.DeleteIndex();
                locked = form.StatusText.StartsWith("İndeks silinemedi") && File.Exists(db);
            }
            form.DeleteIndex();
            bool deleted = !Directory.Exists(index) && !Directory.Exists(att) && form.StatusText.StartsWith("İndeks silindi");
            bool kept = File.Exists(settings.ConfigFile) && File.Exists(searches), cleared = form.ResultCount == 0, enabled = btn != null && btn.Enabled;
            File.Delete(settings.ConfigFile); File.Delete(searches);
            bool ok = enabled && locked && deleted && kept && cleared;
            Console.WriteLine("index_delete button={0} locked={1} deleted={2} settings_and_searches_kept={3} results_cleared={4} -> {5}",
                              enabled, locked, deleted, kept, cleared, ok ? "ok" : "FAIL");
            return ok;
        }

        // "Klasörleri hariç tut" rules (no window: runs on mono on the Mac too): the default names, subfolders, a saved choice
        // in place of the defaults, the settings.json round trip, wrong types, a write error. Works in a new folder under the
        // test home (HATIRMAIL_HOME points there for this check) and deletes it at the end. Prints flags only.
        static bool ExcludeCheck(Settings settings)
        {
            string dir = Path.Combine(settings.Home, "exclude-check-" + Guid.NewGuid().ToString("N"));
            string oldHome = Environment.GetEnvironmentVariable("HATIRMAIL_HOME");
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("exclude_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            Func<string, string, FolderInfo> fi = (id, path) => new FolderInfo { Id = id, Path = path };
            var all = new List<FolderInfo> {
                fi("IN", "\\Gelen Kutusu"), fi("P", "\\Gelen Kutusu\\Özel"), fi("PA", "\\Gelen Kutusu\\Özel\\Alt"), fi("P2", "\\Gelen Kutusu\\Özel2"),
                fi("SENT", "\\Gönderilmiş Öğeler"), fi("DEL", "\\Silinmiş Öğeler"), fi("DELSUB", "\\Silinmiş Öğeler\\Eski proje"),
                fi("JUNK", "\\Önemsiz E-Posta"), fi("JUNK2", "\\Junk Email"), fi("OUT", "\\Giden Kutusu"), fi("DRAFT", "\\Taslaklar"),
                fi("DRAFT2", "\\Drafts"), fi("SYNC", "\\Sync Issues\\Conflicts"), fi("ESIT", "\\Eşitleme Sorunları"), fi("DELE", "\\Deleted Items") };
            Func<HashSet<string>, string> ids = h => string.Join(",", all.Where(x => h.Contains(x.Id)).Select(x => x.Id));
            try
            {
                Directory.CreateDirectory(dir);
                Environment.SetEnvironmentVariable("HATIRMAIL_HOME", dir);
                var s = Settings.Load(); // no settings.json: no choice for any store
                check("defaults", ids(s.ExcludedIn("S", all)) == "DEL,DELSUB,JUNK,JUNK2,OUT,DRAFT,DRAFT2,SYNC,ESIT,DELE");
                s.ExcludedFolders["S"] = new List<string> { "P", "gone" };
                // Özel and its subfolder, not Özel2; the choice replaces the defaults; an unknown id does nothing.
                check("subfolders", ids(s.ExcludedIn("S", all)) == "P,PA");
                check("other_store_defaults", s.ExcludedIn("T", all).Contains("DEL") && !s.ExcludedIn("T", all).Contains("P"));
                s.ExcludedFolders["S"] = new List<string>();
                check("empty_choice", s.ExcludedIn("S", all).Count == 0);
                string sid = "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000";
                s.ExcludedFolders.Clear();
                s.ExcludedFolders[sid] = new List<string> { "00000000AB12", "00000000CD34" };
                s.ExcludedFolders["S2"] = new List<string>();
                s.Threads = 6;
                s.Save();
                var r = Settings.Load();
                check("roundtrip", r.ExcludedFolders.Count == 2 && r.ExcludedFolders[sid].SequenceEqual(s.ExcludedFolders[sid]) &&
                                   r.ExcludedFolders["S2"].Count == 0 && r.Threads == 6 && !File.Exists(s.ConfigFile + ".tmp"));
                File.WriteAllText(s.ConfigFile, "{\"threads\": 3, \"excluded_folders\": {\"A\": \"x\", \"B\": [1, \"f\", null, \"f\"], \"C\": []}}");
                r = Settings.Load();
                check("wrong_types", r.Threads == 3 && !r.ExcludedFolders.ContainsKey("A") && r.ExcludedFolders["B"].SequenceEqual(new[] { "f" }) &&
                                     r.ExcludedFolders["C"].Count == 0);
                File.WriteAllText(s.ConfigFile, "{\"threads\": 3, \"excluded_folders\": [\"x\"], \"keep_alive_min\": 5}");
                r = Settings.Load();
                check("wrong_container", r.ExcludedFolders.Count == 0 && r.Threads == 3 && r.KeepAliveMin == 5);
                File.WriteAllText(s.ConfigFile, "{\"threads\": 4}"); // an older install: the default names apply
                r = Settings.Load();
                check("old_file_defaults", r.ExcludedFolders.Count == 0 && r.ExcludedIn("S", all).Contains("JUNK") &&
                                           r.KeepAfterCloseMin == 5 && r.PauseOnBattery);
                // The after-close time and the battery pause: saved, read back, clamped (a hand edit).
                r.KeepAfterCloseMin = 0; r.PauseOnBattery = false; r.Save();
                var r2 = Settings.Load();
                File.WriteAllText(s.ConfigFile, "{\"keep_after_close_min\": 99}");
                var r3 = Settings.Load();
                File.WriteAllText(s.ConfigFile, "{\"keep_after_close_min\": -4, \"pause_on_battery\": true}");
                var r4 = Settings.Load();
                check("close_battery_settings", r2.KeepAfterCloseMin == 0 && !r2.PauseOnBattery && r3.KeepAfterCloseMin == 10 &&
                                                r4.KeepAfterCloseMin == 0 && r4.PauseOnBattery);
                // A write that fails (the file name is a folder) throws and leaves no temp file.
                File.Delete(s.ConfigFile); Directory.CreateDirectory(s.ConfigFile);
                bool threw = false;
                try { s.Save(); } catch (IOException) { threw = true; } catch (UnauthorizedAccessException) { threw = true; }
                check("write_error", threw && !File.Exists(s.ConfigFile + ".tmp"));
            }
            catch (Exception ex) { Console.WriteLine("exclude_check exception " + ex.GetType().Name + ": " + ex.Message); ok = false; }
            finally
            {
                Environment.SetEnvironmentVariable("HATIRMAIL_HOME", oldHome);
                if (Directory.Exists(dir)) Directory.Delete(dir, true); // only the folder made above
            }
            Console.WriteLine(ok ? "PASS" : "FAIL");
            return ok;
        }

        // Ayarlar "Klasörleri hariç tut" in the window (Windows only: WinForms). The dialog at 2x (layout + screenshot): a
        // checked folder checks its subfolders, a subfolder of a checked folder stays checked, Kaydet gives the top folder.
        // SetExcluded (the code behind the dialog's Kaydet): the İndeks list loses the folder and its subfolders,
        // settings.json has the choice, the results are cleared; a write error keeps the old choice. No worker here (empty
        // home): the purge itself is checked by the worker unit tests and by "--exclude". The start state comes back.
        static bool ExcludeUiCheck(SearchForm form, Settings settings, string storeId, string shotDir)
        {
            bool ok = true;
            Action<string, bool> check = (name, pass) => { Console.WriteLine("exclude_ui_" + name + "=" + (pass ? "ok" : "FAIL")); ok &= pass; };
            var all = form.AllFolders;
            int inbox = all.FindIndex(f => f.Id == "F-INBOX"), proj = all.FindIndex(f => f.Id == "F-PROJ");
            check("start", form.ListedFolderIds().Count == all.Count && form.ExcludedFolderIds.Count == 0 && form.ExcludeButtonEnabled && inbox >= 0 && proj >= 0);
            using (var d = new SearchForm.ExcludeDialog(form, "Sentetik Posta (Atlas Örnek)", all, new HashSet<string>())
                           { StartPosition = FormStartPosition.Manual, Location = new Point(-32000, -32000) })
            {
                d.Show(); Application.DoEvents(); // no owner: see HistoryUiCheck
                // Body, not the dialog: the desktop of the SSH session on 2.19 is small, and Windows made the 2x dialog
                // smaller (client 1030x755 for a 1040x832 body). ponytail: the dialog size itself is not checked here.
                bool fits = Check(d.Body, "exclude-dialog") & A11yCheck(d, "exclude-dialog");
                d.Box.SetItemChecked(inbox, true);
                bool sub = d.Box.GetItemChecked(proj);
                d.Box.SetItemChecked(proj, false); // refused: the checked parent excludes it
                bool kept = d.Box.GetItemChecked(proj);
                var chosen = d.Chosen();
                if (shotDir != null) Shot(d.Body, Path.Combine(shotDir, "ui-exclude-dialog.png"));
                d.Box.SetItemChecked(inbox, false);
                bool off = !d.Box.GetItemChecked(proj);
                check("dialog", fits && sub && kept && chosen.SequenceEqual(new[] { "F-INBOX" }) && off);
                d.Close();
            }
            form.Feed("search", Results()); // results of the old scope: cleared by the change
            bool set = form.SetExcluded(new List<string> { "F-INBOX" });
            var listed = form.ListedFolderIds();
            List<string> v;
            check("apply", set && !listed.Contains("F-INBOX") && !listed.Contains("F-PROJ") && listed.Count == all.Count - 2 && form.ResultCount == 0 &&
                           Settings.Load().ExcludedFolders.TryGetValue(storeId, out v) && v.SequenceEqual(new[] { "F-INBOX" }) &&
                           form.StatusText.StartsWith("Klasör seçimi kaydedildi") && !File.Exists(settings.ConfigFile + ".tmp"));
            File.Delete(settings.ConfigFile); Directory.CreateDirectory(settings.ConfigFile);
            bool failed = !form.SetExcluded(new List<string>()) && form.StatusText.StartsWith("Klasör seçimi kaydedilemedi") &&
                          form.ExcludedFolderIds.Contains("F-INBOX") && !form.ListedFolderIds().Contains("F-INBOX") && !File.Exists(settings.ConfigFile + ".tmp");
            Directory.Delete(settings.ConfigFile);
            check("write_error", failed);
            check("restore", form.SetExcluded(new List<string>()) && form.ListedFolderIds().Count == all.Count && form.ExcludedFolderIds.Count == 0);
            settings.ExcludedFolders.Remove(storeId);
            if (File.Exists(settings.ConfigFile)) File.Delete(settings.ConfigFile);
            return ok;
        }

        // ------------------------------------------------------------------ accessibility checks (Windows only: WinForms)
        /// <summary>An interactive control: it takes the keyboard focus or a click (the inner controls of a NumericUpDown,
        /// DateTimePicker and the hosted controls of a SearchList popup are parts of their owner).</summary>
        static bool Interactive(Control c)
        {
            if (c.Parent is UpDownBase) return false;
            return c is ButtonBase || c is TextBoxBase || c is ComboBox || c is ListControl || c is ListView || c is DateTimePicker ||
                   c is UpDownBase || c is TabControl || c is TabPage || c is TrackBar;
        }

        /// <summary>The accessibility audit of a window part: each visible interactive control has a name (no "&amp;") and a
        /// description that is not the name; TabIndex values of the children of each container are unique; the Tab order
        /// follows the screen (top to bottom, then left to right in one row); every tab of a TabControl can be reached
        /// (the TabControl is a tab stop, each page has a tab stop); the access keys of one tab (and the tab keys) are
        /// unique; every color is a system color. Prints one line per rule.</summary>
        public static bool A11yCheck(Control root, string name)
        {
            bool ok = true;
            Action<string, bool, string> check = (rule, pass, detail) =>
            {
                Console.WriteLine("a11y_{0} {1}{2} -> {3}", rule, name, pass || detail.Length == 0 ? "" : " " + detail, pass ? "ok" : "FAIL");
                ok &= pass;
            };
            // Called once for each tab (the controls of the hidden tabs are not Visible); the tab pages themselves always.
            var all = new[] { root }.Concat(SearchForm.AllControls(root)).ToList();
            var pages = all.OfType<TabPage>().ToList();
            var named = all.Where(c => c.Visible && Interactive(c)).Concat(pages).Distinct().ToList();
            var noName = named.Where(c => string.IsNullOrWhiteSpace(c.AccessibleName) || c.AccessibleName.Contains("&")).ToList();
            check("names", noName.Count == 0, string.Join(", ", noName.Select(Who)));
            var noDesc = named.Where(c => string.IsNullOrWhiteSpace(c.AccessibleDescription) || c.AccessibleDescription == c.AccessibleName).ToList();
            check("descriptions", noDesc.Count == 0, string.Join(", ", noDesc.Select(Who)));
            var clash = all.Where(c => c.Controls.Count > 1 && !(c is UpDownBase) && !(c is TabControl))
                           .Where(c => c.Controls.Cast<Control>().GroupBy(x => x.TabIndex).Any(g => g.Count() > 1)).ToList();
            check("tabindex_unique", clash.Count == 0, string.Join(", ", clash.Select(c => Who(c) + "[" + string.Join(",", c.Controls.Cast<Control>().Select(x => x.TabIndex)) + "]")));
            // Tab order: walk GetNextControl from the root through the visible, enabled tab stops.
            var order = new List<Control>();
            for (var c = root.GetNextControl(null, true); c != null && order.Count < 2000; c = root.GetNextControl(c, true))
            {
                var x = c.Parent is UpDownBase ? c.Parent : c; // the edit field of a NumericUpDown = the NumericUpDown
                if (c.CanSelect && c.TabStop && (order.Count == 0 || order[order.Count - 1] != x)) order.Add(x);
            }
            var bad = new List<string>();
            for (int i = 1; i < order.Count; i++)
            {
                Rectangle a = At(root, order[i - 1]), b = At(root, order[i]);
                bool sameTab = PageOf(order[i - 1]) == PageOf(order[i]);
                // Next in reading order: lower, or in the same row (vertical overlap) and to the right. 4 px tolerance.
                bool row = b.Top < a.Bottom - 4 && a.Top < b.Bottom - 4;
                if (sameTab && !(b.Top >= a.Top - 4 && !row || row && b.Left >= a.Left - 4)) bad.Add(Who(order[i - 1]) + " > " + Who(order[i]));
            }
            check("tab_order_visual", bad.Count == 0 && order.Count > 0, string.Join("; ", bad));
            var tcs = all.OfType<TabControl>().ToList();
            var unreachable = tcs.Where(t => !t.TabStop).Select(Who)
                .Concat(pages.Where(p => p.Visible && !SearchForm.AllControls(p).Any(c => c.CanSelect && c.TabStop)).Select(Who)).ToList();
            check("tabs_reachable", unreachable.Count == 0, string.Join(", ", unreachable));
            // Access keys: per tab page (with the window-level labels and buttons outside the tabs, and the tab keys).
            var keyBad = new List<string>();
            var outside = all.Where(c => (c is Label || c is ButtonBase) && !pages.Any(p => p.Contains(c))).ToList();
            foreach (var scope in pages.Count > 0 ? pages.Select(p => (Control)p) : new Control[] { null })
            {
                var owners = (scope == null ? Enumerable.Empty<Control>() : SearchForm.AllControls(scope)).Concat(outside).Where(c => c is Label && ((Label)c).UseMnemonic || c is ButtonBase)
                    .Select(c => new { C = c, K = char.ToUpper(SearchForm.Mnemonic(c.Text), System.Globalization.CultureInfo.GetCultureInfo("tr-TR")) })
                    .Where(x => x.K != '\0').ToList();
                var keys = owners.Select(x => x.K.ToString()).ToList();
                if (pages.Count > 0) keys.AddRange(new[] { "A", "İ", "Y" });
                foreach (var g in keys.GroupBy(x => x).Where(g => g.Count() > 1)) keyBad.Add((scope == null ? "" : Who(scope)) + ":Alt+" + g.Key);
            }
            check("mnemonics_unique", keyBad.Count == 0, string.Join(", ", keyBad));
            var colors = all.Where(c => c.Visible)
                            .SelectMany(c => new[] { Tuple.Create(c, c.ForeColor), Tuple.Create(c, c.BackColor) })
                            .Where(t => !t.Item2.IsSystemColor && t.Item2 != Color.Transparent && t.Item2.A != 0).ToList();
            check("color_system_only", colors.Count == 0, string.Join(", ", colors.Take(8).Select(t => Who(t.Item1) + "=" + t.Item2.Name)));
            return ok;
        }

        static TabPage PageOf(Control c) { for (var x = c; x != null; x = x.Parent) if (x is TabPage) return (TabPage)x; return null; }

        static Rectangle At(Control root, Control c) { return root.RectangleToClient(c.Parent.RectangleToScreen(c.Bounds)); }

        static string Who(Control c)
        {
            string n = !string.IsNullOrEmpty(c.AccessibleName) ? c.AccessibleName : !string.IsNullOrEmpty(c.Text) ? SearchForm.StripAmp(c.Text) : c.Name;
            return c.GetType().Name + "'" + Short(n ?? "") + "'";
        }

        [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [MarshalAs(UnmanagedType.IUnknown)] out object acc);

        /// <summary>The MSAA name (what a screen reader gets) of a window or of child id of it (a list row: 1-based).</summary>
        static string MsaaName(Control c, int child)
        {
            var iid = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71"); // IAccessible
            object o;
            if (AccessibleObjectFromWindow(c.Handle, 0xFFFFFFFC, ref iid, out o) != 0 || o == null) return null; // OBJID_CLIENT
            try { return ((Accessibility.IAccessible)o).get_accName(child); }
            catch (Exception) { return null; }
            finally { Marshal.ReleaseComObject(o); }
        }

        /// <summary>Names as MSAA gives them to a screen reader: the search box, a filter list, a date field, a button, the
        /// tabs, and the first result row (the sentence).</summary>
        static bool MsaaCheck(SearchForm form)
        {
            var all = SearchForm.AllControls(form).ToList();
            var want = new[] { form.QueryBox, form.FilterBoxes[0], (Control)form.DatePickers[0],
                               all.OfType<Button>().First(b => b.AccessibleName == "Ara"), form.Tabs };
            var bad = want.Where(c => MsaaName(c, 0) != c.AccessibleName).Select(c => Who(c) + "=" + (MsaaName(c, 0) ?? "null")).ToList();
            bool row = form.ResultCount > 0 && MsaaName(form.Results, 1) == form.ResultSentence(0);
            Console.WriteLine("a11y_msaa_names {0} -> {1}", string.Join(", ", bad), bad.Count == 0 ? "ok" : "FAIL");
            Console.WriteLine("row_text_msaa first='{0}' -> {1}", form.ResultCount > 0 ? MsaaName(form.Results, 1) : "-", row ? "ok" : "FAIL");
            return bad.Count == 0 && row;
        }

        /// <summary>Result rows: the native text is the sentence and the other columns are empty; the cells keep the column
        /// texts (sort, layout). High contrast while the window is open: the marks are bold and underlined, all colors stay
        /// system colors. The focus frame: around the search box, a filter list, the result list, a tab header.</summary>
        static bool ColorRowRingCheck(SearchForm form, string shotDir)
        {
            bool ok = true;
            Action<string, bool> check = (n, pass) => { Console.WriteLine(n + " -> " + (pass ? "ok" : "FAIL")); ok &= pass; };
            form.ShowTab(0);
            form.Feed("search", Results());
            var lv = form.Results;
            bool native = lv.Items.Count == 3 && Enumerable.Range(0, 3).All(i => lv.Items[i].Text == form.ResultSentence(i) &&
                          lv.Items[i].SubItems.Cast<ListViewItem.ListViewSubItem>().Skip(1).All(x => x.Text == ""));
            string s0 = form.ResultSentence(0);
            check("row_text_native first='" + s0 + "'", native && s0.StartsWith("Ahmet Demir, Kargo faturası – Eylül 2026 (Hızlı Lojistik A.Ş.), 12 Eylül") &&
                                                    s0.Contains(", Gelen Kutusu, ekte eşleşti: Hızlı Lojistik") && s0.EndsWith("(Kargo_Faturalari_2026.zip içinde)") &&
                                                    form.ResultSentence(1).Contains(", Projeler") && form.ResultSentence(2).EndsWith("zayıf eşleşme") &&
                                                    form.ResultCells(0)[1] == "Ahmet Demir");
            // High contrast on (simulated): Windows sends WM_SYSCOLORCHANGE; the window follows without a restart.
            SearchForm.HighContrastOverride = true;
            form.SimulateSystemColorsChanged();
            form.SelectResult(0); Application.DoEvents();
            int marked = form.PreviewMarkedChars;
            bool hcColors = A11yCheck(form.Root, "high-contrast");
            if (shotDir != null) Shot(form.Root, Path.Combine(shotDir, "ui-high-contrast.png"));
            SearchForm.HighContrastOverride = false;
            form.SimulateSystemColorsChanged();
            form.SelectResult(0); Application.DoEvents();
            check("hc_live marked=" + marked, marked > 0 && hcColors && form.PreviewMarkedChars == marked);
            SearchForm.HighContrastOverride = null;
            form.SimulateSystemColorsChanged();
            // Focus frame: outside the target on all sides, inside the window.
            foreach (var t in new Control[] { form.QueryBox, form.FilterBoxes[0], lv, form.Tabs })
            {
                form.SetActive(t); Application.DoEvents();
                var r = form.RingBounds;
                var tr = t == form.Tabs ? form.Root.RectangleToClient(t.RectangleToScreen(form.Tabs.GetTabRect(0))) : At(form.Root, t);
                var want = Rectangle.Inflate(tr, 2, 2);
                want.Intersect(new Rectangle(Point.Empty, form.Root.Size));
                bool around = r.Width > 0 && r.Contains(want);
                check("ring_" + t.AccessibleName + " ring=" + r + " target=" + tr, around);
            }
            form.SetActive(lv); Application.DoEvents();
            if (shotDir != null) Shot(form.Root, Path.Combine(shotDir, "ui-focus-list.png"));
            return ok;
        }

        /// <summary>Keyboard: focus at the window open, Esc back to the search box, Ctrl+F, Enter opens, Shift+F10 and the
        /// menu key open the result menu (none without a selected row), Ctrl+1..5 sort, Alt+tab keys, Alt+N, F1 list, the
        /// first-result setting.</summary>
        static bool KeyCheck(SearchForm form, JsonSource src, string shotDir)
        {
            bool ok = true;
            Action<string, bool> check = (n, pass) => { Console.WriteLine("key_" + n + " -> " + (pass ? "ok" : "FAIL")); ok &= pass; };
            form.ShowTab(0);
            form.Feed("search", Results());
            var lv = form.Results;
            form.SetActive(lv); form.ShowTab(1); form.SetActive(form.Tabs);
            form.PressKey(Keys.Escape);
            bool esc1 = form.Active == form.QueryBox && form.Tabs.SelectedIndex == 0;
            form.SetActive(lv);
            form.PressKey(Keys.Escape);
            check("esc_to_query", esc1 && form.Active == form.QueryBox);
            form.ShowTab(2); form.SetActive(form.Tabs);
            form.PressKey(Keys.Control | Keys.F);
            check("ctrl_f", form.Active == form.QueryBox && form.Tabs.SelectedIndex == 0 && ((TextBox)form.QueryBox).SelectionLength == form.QueryBox.Text.Length);
            // Enter in the list (KeyDown, as the key message gives it to the list).
            form.SelectResult(1); form.SetActive(lv);
            int before = src.Opened.Count;
            var kd = typeof(Control).GetMethod("OnKeyDown", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            kd.Invoke(lv, new object[] { new KeyEventArgs(Keys.Enter) });
            check("enter_open", src.Opened.Count == before + 1 && src.Opened.Last() == form.SelectedEid);
            // Shift+F10 / menu key = WM_CONTEXTMENU with lParam -1.
            SendMessage(lv.Handle, 0x7B, lv.Handle, (IntPtr)(-1)); Application.DoEvents();
            bool menu = form.ResultMenu.Visible;
            form.ResultMenu.Close(); Application.DoEvents();
            lv.SelectedItems.Clear();
            SendMessage(lv.Handle, 0x7B, lv.Handle, (IntPtr)(-1)); Application.DoEvents();
            bool none = !form.ResultMenu.Visible;
            form.ResultMenu.Close();
            check("shift_f10 menu=" + menu + " no_row_no_menu=" + none, menu && none);
            form.SelectResult(0);
            // Ctrl+1..5 (and the number pad): sort by the column; again: the other order.
            var heads = new List<string>();
            for (int i = 0; i < 5; i++) { form.PressKey(Keys.Control | (Keys.D1 + i)); heads.Add(form.HeaderText(i)); }
            form.PressKey(Keys.Control | Keys.NumPad3); form.PressKey(Keys.Control | Keys.NumPad3); // the number pad; twice: the other order
            bool again = form.HeaderText(2) == "Tarih ▲";
            form.PressKey(Keys.Control | Keys.D4);
            check("ctrl_sort " + string.Join("|", heads), heads.SequenceEqual(new[] { "Konu ▲", "Kişi ▲", "Tarih ▼", "Güven ▼", "Neden ▲" }) && again &&
                                                       form.HeaderText(3) == "Güven ▼");
            if (form.HeaderText(3) != "Güven ▼") form.ClickHeader(3);
            // Access keys: the tabs (Alt+A, Alt+İ, Alt+Y; the TabControl shows "&" as text) and a filter label.
            bool tabsOk = true;
            foreach (var t in new[] { Tuple.Create('i', 1), Tuple.Create('Y', 2), Tuple.Create('a', 0) })
            {
                form.PressMnemonic(t.Item1);
                tabsOk &= form.Tabs.SelectedIndex == t.Item2 && form.Active == form.Tabs;
            }
            check("mnemonic_tabs", tabsOk);
            form.SetActive(form.QueryBox);
            form.PressMnemonic('n');
            check("mnemonic_field active=" + (form.Active != null ? form.Active.AccessibleName : "-"), form.Active == form.FilterBoxes[0]);
            // Second ways (a host loop without pre-translation): Alt+letter as WM_SYSCOMMAND SC_KEYMENU; a key message that
            // reaches the focused control (WM_KEYDOWN Esc to the list: KeyPreview).
            form.SetActive(form.QueryBox);
            SendMessage(form.Handle, 0x112, (IntPtr)0xF100, (IntPtr)'l');
            bool sys = form.Active == form.FilterBoxes[1];
            form.SetActive(lv);
            SendMessage(lv.Handle, 0x100, (IntPtr)0x1B, IntPtr.Zero);
            check("fallback sc_keymenu=" + sys + " keypreview_esc=" + (form.Active == form.QueryBox), sys && form.Active == form.QueryBox);
            // F1: the list has every shortcut and every access key; the dialog passes the layout and audit checks.
            string text = null; bool dlgOk = false;
            form.ShowModal = d =>
            {
                d.StartPosition = FormStartPosition.Manual; d.Location = new Point(-32000, -32000);
                d.Show(); Application.DoEvents();
                var sd = (SearchForm.ShortcutDialog)d;
                text = sd.Box.Text;
                dlgOk = Check(sd.Body, "shortcut-dialog") & A11yCheck(sd, "shortcut-dialog") && sd.CancelButton != null; // Body: see ExcludeUiCheck
                if (shotDir != null) Shot(sd.Body, Path.Combine(shotDir, "ui-shortcuts.png"));
                d.Close();
            };
            form.PressKey(Keys.F1);
            form.ShowModal = null;
            var keysOnTabs = new[] { "Alt+R: Ara", "Alt+U: Aramayı durdur", "Alt+N: Gönderen", "Alt+Ş: Bitiş", "Alt+P: Posta kutusu",
                                     "Alt+K: Klasörler", "Alt+B: Yazı büyüklüğü", "Alt+L: Arama bitince ilk sonuca geç" };
            bool listed = text != null && SearchForm.Shortcuts.All(sc => text.Contains(sc[0] + ": " + sc[1])) && keysOnTabs.All(text.Contains);
            check("f1 lines=" + (text ?? "").Split('\n').Length, listed && dlgOk);
            // Window open: the focus is in the search box (Load).
            // "Arama bitince ilk sonuca geç": the final list moves the focus only with the setting on, while the focus is in
            // the query row and the text did not change.
            form.FocusFirstSetting = true;
            lv.SelectedItems.Clear(); // no row chosen by the user: the first row of the new list is selected
            form.FocusQuery(); form.MarkSearchStarted();
            form.Feed("search", Results());
            bool moved = form.Active == lv && lv.FocusedItem != null && lv.FocusedItem.Index == 0 && lv.Items[0].Selected;
            form.FocusQuery(); form.MarkSearchStarted(); form.SetQuery(form.QueryBox.Text + " x");
            form.Feed("search", Results());
            bool textChanged = form.Active == form.QueryBox;
            form.FocusFirstSetting = false;
            form.FocusQuery(); form.MarkSearchStarted();
            form.Feed("search", Results());
            bool off = form.Active == form.QueryBox;
            check("first_result moved=" + moved + " text_changed_stays=" + textChanged + " setting_off_stays=" + off, moved && textChanged && off);
            form.SetQuery("geçen ay Ahmet'ten gelen kargo faturası");
            return ok;
        }


        /// <summary>Text size 150 % and 200 % (Ayarlar, Kaydet): every tab at the default and minimum size passes the
        /// layout, filter grid and audit checks; back to 100 % the sizes are the same as before (±2 px).</summary>
        static bool ScaleCheck(SearchForm form, Settings settings, string shotDir)
        {
            bool ok = true;
            var root = form.Root;
            var all = SearchForm.AllControls(form).ToList();
            var sizeBox = all.OfType<ComboBox>().First(x => x.AccessibleName == "Yazı büyüklüğü");
            var save = all.OfType<Button>().First(x => x.Text == "&Kaydet");
            Func<int[]> dims = () => new[] { form.QueryBox.Height, form.FilterBoxes[0].Height, form.Results.Columns[2].Width, form.Font.Height };
            ShowAllFilterRows(form);
            var start = dims();
            Directory.CreateDirectory(settings.Home);
            foreach (int pct in new[] { 150, 200, 100 })
            {
                form.ShowTab(2);
                sizeBox.SelectedIndex = (pct - 100) / 25;
                save.PerformClick();
                Application.DoEvents();
                bool applied = form.TextScalePct == pct && Settings.Load().TextScalePct == pct;
                Console.WriteLine("scale_text_{0} applied={1} font={2:0.0}pt -> {3}", pct, applied, form.Font.SizeInPoints, applied ? "ok" : "FAIL");
                ok &= applied;
                if (pct == 100) break;
                foreach (var size in new[] { form.DesignClientSize, form.MinClientSize })
                {
                    root.Bounds = new Rectangle(Point.Empty, size);
                    for (int tab = 0; tab < 3; tab++)
                    {
                        form.ShowTab(tab); root.PerformLayout(); Application.DoEvents();
                        string n = "text" + pct + "-" + (size == form.DesignClientSize ? "default" : "min") + "-" + new[] { "ara", "indeks", "ayarlar" }[tab];
                        ok &= Check(root, n);
                        if (tab == 0) { ok &= FilterGrid(form, n); ok &= A11yCheck(root, n); }
                        if (shotDir != null && tab == 0 && size == form.DesignClientSize) Shot(root, Path.Combine(shotDir, "ui-text" + pct + "-ara.png"));
                    }
                }
            }
            root.Bounds = new Rectangle(Point.Empty, form.DesignClientSize); root.PerformLayout(); Application.DoEvents();
            var end = dims();
            bool back = start.Zip(end, (a, b) => Math.Abs(a - b) <= 2).All(x => x);
            Console.WriteLine("scale_text_back start={0} end={1} -> {2}", string.Join("/", start), string.Join("/", end), back ? "ok" : "FAIL");
            if (File.Exists(settings.ConfigFile)) File.Delete(settings.ConfigFile);
            form.ClickClearFilters();
            return ok && back;
        }

        /// <summary>A small screen: the window at 200 % text in 1366x690 (and 2x DPI at 125 % text): the pages scroll, no
        /// control is cut (Check on the scrolled page area). Then a DPI change 192 -> 288 -> 192 (WM_DPICHANGED as Windows
        /// sends it): each step passes the layout checks, and the sizes come back (±2 px).</summary>
        static bool DpiCheck(SearchForm form, string shotDir)
        {
            bool ok = true;
            var root = form.Root;
            Func<int[]> dims = () => new[] { form.QueryBox.Height, form.FilterBoxes[0].Height, form.Results.Columns[2].Width, form.Font.Height };
            ShowAllFilterRows(form);
            var start = dims();
            int d0 = form.WindowDpi;
            foreach (int to in new[] { d0 * 3 / 2, d0 })
            {
                var wr = form.Bounds;
                var rc = new int[] { wr.Left, wr.Top, wr.Left + wr.Width * to / form.WindowDpi, wr.Top + wr.Height * to / form.WindowDpi };
                var h = GCHandle.Alloc(rc, GCHandleType.Pinned);
                try { SendMessage(form.Handle, 0x02E0, (IntPtr)(to << 16 | to), h.AddrOfPinnedObject()); }
                finally { h.Free(); }
                Application.DoEvents();
                var size = new Size(form.DesignClientSize.Width, form.DesignClientSize.Height);
                root.Bounds = new Rectangle(Point.Empty, size);
                bool step = form.WindowDpi == to;
                for (int tab = 0; tab < 3; tab++)
                {
                    form.ShowTab(tab); root.PerformLayout(); Application.DoEvents();
                    step &= Check(root, "dpi" + to + "-" + new[] { "ara", "indeks", "ayarlar" }[tab]);
                    if (tab == 0) step &= FilterGrid(form, "dpi" + to);
                }
                if (shotDir != null && to != d0) Shot(root, Path.Combine(shotDir, "ui-dpi" + to + "-ara.png"));
                Console.WriteLine("dpi_change to={0} window_dpi={1} -> {2}", to, form.WindowDpi, step ? "ok" : "FAIL");
                ok &= step;
            }
            var end = dims();
            bool back = start.Zip(end, (a, b) => Math.Abs(a - b) <= 2).All(x => x);
            Console.WriteLine("dpi_change_back start={0} end={1} -> {2}", string.Join("/", start), string.Join("/", end), back ? "ok" : "FAIL");
            // Small screen: the real window of this check is limited to the 1024x768 screen of the SSH session, smaller
            // than the minimum window at 2x (1640x1120). The root fills the window as in Outlook (Dock Fill): the window
            // scrolls, the root keeps the minimum size, and every tab passes the layout checks at that size.
            root.Dock = DockStyle.Fill;
            form.PerformLayout(); Application.DoEvents();
            bool small = true;
            for (int tab = 0; tab < 3; tab++)
            {
                form.ShowTab(tab); form.PerformLayout(); Application.DoEvents();
                bool scrolls = form.HorizontalScroll.Visible && form.VerticalScroll.Visible && form.ClientSize.Width < form.AutoScrollMinSize.Width;
                bool size = root.Width >= form.AutoScrollMinSize.Width && root.Height >= form.AutoScrollMinSize.Height;
                bool inside = Check(root, "small-" + new[] { "ara", "indeks", "ayarlar" }[tab]);
                Console.WriteLine("scale_small_screen tab={0} window_client={1} root={2} min={3} scrolls={4} -> {5}", tab, form.ClientSize, root.Size,
                                  form.AutoScrollMinSize, scrolls, scrolls && size && inside ? "ok" : "FAIL");
                small &= scrolls && size && inside;
            }
            if (shotDir != null) { form.ShowTab(0); Shot(root, Path.Combine(shotDir, "ui-small-screen.png")); }
            root.Dock = DockStyle.None;
            root.Bounds = new Rectangle(Point.Empty, form.DesignClientSize); root.PerformLayout(); Application.DoEvents();
            // Per-monitor DPI awareness of a window made inside Native.DpiScope (as Connect.ShowForm makes the window).
            int aware;
            using (new Native.DpiScope())
            using (var f = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual, Location = new Point(-32000, -32000) })
            { f.Show(); aware = Native.WindowAwareness(f.Handle); f.Close(); }
            Console.WriteLine("dpi_aware awareness={0} -> {1}", aware, aware == 2 ? "ok" : "FAIL");
            form.ClickClearFilters();
            return ok && back && small && aware == 2;
        }

        /// <summary>Results, filter lists, "Özel aralık…" (the third filter row) and a query: the fullest Ara tab.</summary>
        static void ShowAllFilterRows(SearchForm form)
        {
            form.Feed("search", Results());
            form.Feed("status", Facets(false));
            form.FilterBoxes[3].SelectedIndex = form.FilterBoxes[3].Items.Count - 1; form.FilterBoxes[4].SelectedIndex = 3;
        }

        // Result row sentences (no window: runs on mono on the Mac too).
        static bool RowTextCheck()
        {
            bool ok = true;
            var today = new DateTime(2026, 10, 26);
            Func<string, string, string, string, Dictionary<string, object>> r = (sender, subject, date, conf) =>
                new Dictionary<string, object> { {"sender", sender}, {"subject", subject}, {"date", date}, {"confidence", conf}, {"match_attachment", null} };
            var full = r("Ayşe Yılmaz", "Bütçe taslağı", "2026-03-12T09:30:00", "yüksek"); full["match_attachment"] = "rapor.xlsx";
            var zip = r("Ahmet Demir", "Kargo", "2026-09-01T08:00:00", "orta"); zip["match_attachment"] = "Arsiv.zip"; zip["match_att_member"] = "Eylül\\fatura.pdf";
            var cases = new[] {
                Tuple.Create("full", RowText.Sentence(full, "\\Gelen Kutusu", today), "Ayşe Yılmaz, Bütçe taslağı, 12 Mart, Gelen Kutusu, ekte eşleşti: rapor.xlsx"),
                Tuple.Create("no_attachment", RowText.Sentence(r("Ayşe Yılmaz", "Bütçe", "2026-10-01T00:00:00", "orta"), "\\Gelen Kutusu\\Projeler", today), "Ayşe Yılmaz, Bütçe, 1 Ekim, Projeler"),
                Tuple.Create("zip_member", RowText.Sentence(zip, "\\Arşiv", today), "Ahmet Demir, Kargo, 1 Eylül, Arşiv, ekte eşleşti: fatura.pdf (Arsiv.zip içinde)"),
                Tuple.Create("weak", RowText.Sentence(r("Can", "Sayım", "2026-08-20T14:05:00", "zayıf"), null, today), "Can, Sayım, 20 Ağustos, zayıf eşleşme"),
                Tuple.Create("empty_parts", RowText.Sentence(r("", "  ", "", "kelime"), "", today), ""),
                Tuple.Create("other_year", RowText.Sentence(r("Zeynep", "Teklif", "2025-12-31T23:00:00", "yüksek"), "\\Gelen Kutusu", today), "Zeynep, Teklif, 31 Aralık 2025, Gelen Kutusu") };
            foreach (var c in cases)
            {
                bool pass = c.Item2 == c.Item3;
                Console.WriteLine("row_text_{0} -> {1}{2}", c.Item1, pass ? "ok" : "FAIL", pass ? "" : " got='" + c.Item2 + "'");
                ok &= pass;
            }
            Console.WriteLine(ok ? "PASS" : "FAIL");
            return ok;
        }

        // Checks on the visible controls: text fits (labels, buttons, check boxes, tabs), controls stay inside the
        // window, list columns fill the list without a cut. A control in a container that scrolls (AutoScroll: the
        // Ayarlar page) must be inside the scrolled area of that container (DisplayRectangle), not inside its visible part.
        static bool Check(Control root, string name)
        {
            bool ok = true;
            var rr = root.RectangleToScreen(root.ClientRectangle);
            int n = 0;
            foreach (var c in SearchForm.AllControls(root).Where(c => c.Visible))
            {
                n++;
                // Inside the client area of its parent (and so inside the window): nothing is clipped. Bounds = the full
                // window of the control (its border included) in the client coordinates of the parent.
                var r = c.Parent.RectangleToScreen(c.Bounds);
                var sp = c.Parent as ScrollableControl;
                bool scrolls = sp != null && sp.AutoScroll;
                // The scroll area = DisplayRectangle with the padding of the page (DisplayRectangle leaves the padding out).
                var pr = c.Parent.RectangleToScreen(scrolls ? Rectangle.FromLTRB(sp.DisplayRectangle.Left - sp.Padding.Left, sp.DisplayRectangle.Top - sp.Padding.Top,
                                                                                 sp.DisplayRectangle.Right + sp.Padding.Right, sp.DisplayRectangle.Bottom + sp.Padding.Bottom)
                                                            : c.Parent.ClientRectangle);
                bool inScroll = false;
                for (var x = c.Parent; x != null && x != root; x = x.Parent) { var xs = x as ScrollableControl; if (xs != null && xs.AutoScroll) inScroll = true; }
                if (r.Width > 0 && r.Height > 0 && (!pr.Contains(r) || !inScroll && !rr.Contains(r)))
                { Console.WriteLine("{0}: outside {1} '{2}' {3} parent={4}", name, c.GetType().Name, Short(c.Text), r, pr); ok = false; }
                var l = c as Label;
                if (l != null && l.Text.Length > 0)
                {
                    // Height for the text wrapped at the label width (one line for a label that ends with "…"),
                    // and the longest word fits the width. A shortcut "&" is not drawn.
                    string lt = l.UseMnemonic ? l.Text.Replace("&&", "\u0001").Replace("&", "").Replace("\u0001", "&") : l.Text;
                    int w = l.ClientSize.Width - l.Padding.Horizontal;
                    int need = (l.AutoEllipsis ? TextRenderer.MeasureText("Ag", l.Font).Height
                                : TextRenderer.MeasureText(lt, l.Font, new Size(Math.Max(1, w), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl).Height)
                               + l.Padding.Vertical;
                    int word = TextRenderer.MeasureText(LongestWord(lt), l.Font).Width;
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
                    // Every cell except the first column (subject, folder: long texts end with "…") fits its column. The
                    // result list draws its cells from the result ("cells"); its row text is the sentence.
                    foreach (ListViewItem it in lv.Items)
                        for (int col = 1; col < lv.Columns.Count; col++)
                        {
                            var d = it.Tag as Dictionary<string, object>;
                            string t = d != null && d.ContainsKey("cells") ? ((string[])d["cells"])[col] : it.SubItems[col].Text;
                            if (TextRenderer.MeasureText(t, lv.Font).Width + 12 > lv.Columns[col].Width)
                            { Console.WriteLine("{0}: cell cut col={1} '{2}' w={3}", name, col, t, lv.Columns[col].Width); ok = false; }
                        }
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
                            TextRenderer.DrawText(g, tb.Text, tb.Font, Rectangle.Inflate(r, -4, tb.Multiline ? -4 : 0), tb.Enabled ? tb.ForeColor : SystemColors.GrayText,
                                                  (tb.Multiline ? TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl
                                                                : TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis) |
                                                  (tb.TextAlign == HorizontalAlignment.Right ? TextFormatFlags.Right : TextFormatFlags.Left));
                        }
                        else if (c is SearchForm.FocusFrame) // DrawToBitmap draws the children in reverse order: the frame is below
                        {
                            int w = ((SearchForm.FocusFrame)c).Thickness;
                            using (var br = new SolidBrush(c.BackColor))
                            {
                                g.FillRectangle(br, r.X, r.Y, r.Width, w); g.FillRectangle(br, r.X, r.Bottom - w, r.Width, w);
                                g.FillRectangle(br, r.X, r.Y, w, r.Height); g.FillRectangle(br, r.Right - w, r.Y, w, r.Height);
                            }
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
            // Dates and senders differ: the sort check (SortUiCheck) orders by them. sender_key of Ahmet Demir = his key in Facets().
            var dates = new Dictionary<string, string> { {"E1", "2026-09-12T09:41:00"}, {"E2", "2026-08-20T14:05:00"}, {"E3", "2026-10-01T08:30:00"} };
            Func<string, string, string, string, string, ArrayList, ArrayList, Dictionary<string, object>> r = (eid, subj, who, conf, snip, hl, why) =>
                new Dictionary<string, object> { {"eid", eid}, {"subject", subj}, {"sender", who}, {"recipients", "Furkan Çakır; Muhasebe Ekibi"},
                    {"sender_key", who == "Ahmet Demir" ? "ahmet.demir@hizli-lojistik.example" : "ayse@x.example"},
                    {"date", dates[eid]}, {"confidence", conf}, {"why", why}, {"snippet", snip}, {"hl", hl},
                    {"folder", eid == "E1" ? "F-INBOX" : eid == "E2" ? "F-PROJ" : "F-ARCH"},
                    {"hl_subject", new ArrayList { new ArrayList { 0, 5 }, new ArrayList { 6, 8 } }}, {"attachments", "Fatura_2026-09.pdf; İrsaliye.xlsx"},
                    {"match_attachment", null} };
            string s1 = "Merhaba Furkan Bey, eylül ayına ait kargo faturası ekte gönderilmiştir. Toplam tutar 4.250 TL olup ödeme vadesi 30 gündür. İyi çalışmalar.";
            var list = new ArrayList {
                r("E1", "Kargo faturası – Eylül 2026 (Hızlı Lojistik A.Ş.)", "Ahmet Demir", "yüksek", s1,
                  new ArrayList { new ArrayList { s1.IndexOf("kargo"), 5 }, new ArrayList { s1.IndexOf("faturası"), 8 } }, new ArrayList { "kelime 2/2", "anlam 0.83", "ekte" }),
                r("E2", "Kargo faturası – Ağustos", "Ahmet Demir", "orta", "Ağustos kargo faturası ve irsaliye ekte.", new ArrayList { new ArrayList { 8, 5 }, new ArrayList { 14, 8 } },
                  new ArrayList { "kelime 2/2", "anlam 0.74", "tarih yakın" }),
                r("E3", "Depo sayım sonuçları", "Ayşe Kaya", "zayıf", "Sayım dosyası ektedir.", new ArrayList(), new ArrayList { "kelime 1/2", "anlam 0.61", "ekte", "tarih dışı" }) };
            // E1: a match in a zip member, with a long path (the bar line ends with "…"). E3: a result of an older worker
            // (attachment name, no index): the line shows, the button is off. E2: a match in the mail text: no bar.
            var e1 = (Dictionary<string, object>)list[0];
            e1["match_attachment"] = "Kargo_Faturalari_2026.zip"; e1["match_att_index"] = 1;
            e1["match_att_member"] = "Eylül/Hızlı Lojistik – Kargo faturası TR-4471 (düzeltilmiş, imzalı ve kaşeli son sürüm).pdf";
            ((Dictionary<string, object>)list[2])["match_attachment"] = "Sayim_Eylul.xlsx";
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

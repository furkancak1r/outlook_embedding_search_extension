// HatırMail — Classic Outlook COM add-in (.NET Framework 4.8, C# 7.3, late binding, no PIA dependency).
//
// Design:
//  - Outlook loads Connect (IDTExtensibility2). It adds a ribbon button and nothing else at startup:
//    no model, no worker, no timer, no polling, no background indexing.
//  - The search window starts ONE worker process (embedded Python) on a user action (window open, Ara, İndeksleme).
//    The add-in reads mail with the Outlook Object Model on the UI thread (read-only properties only)
//    and sends JSON lines over the anonymous stdin/stdout pipes of its own child process. No port, no network.
//  - While the window is open the worker keeps the model loaded for the next search (setting keep_alive_min,
//    0 = exit after each search). Window close, Outlook close, cancel, error, timeout and keep-alive expiry end it
//    (Job Object kill-on-close).
//  - "Outlook'ta Aç" calls Namespace.GetItemFromID(...).Display(). It never sends, deletes, moves or marks mail.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

[assembly: AssemblyTitle("HatirMail")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]
[assembly: ComVisible(false)]

namespace HatirMail
{
    // ---------------------------------------------------------------------------------------------
    // COM interfaces (declared locally so that the build needs no Office PIA / extensibility.dll).
    // ---------------------------------------------------------------------------------------------
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object Application, int ConnectMode,
                                      [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
                                      [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection(int RemoveMode,
                                         [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    // AutoDispatch: Office calls the ribbon callback (OnSearch) by name through IDispatch.
    [ComVisible(true), Guid("2B7C9E14-6A3F-4D85-A1E2-7F90C4B3D215"), ProgId("HatirMail.Connect"),
     ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        object app;
        SearchForm form;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            app = application; // nothing else runs at startup
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try { if (form != null && !form.IsDisposed) form.Shutdown(); } catch { }
            form = null;
            app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom)
        {
            // Outlook closes: the window closes and its worker stops (also a running index; the work is saved).
            try { if (form != null && !form.IsDisposed) form.Shutdown(); } catch { }
        }

        public string GetCustomUI(string ribbonId)
        {
            if (ribbonId != "Microsoft.Outlook.Explorer") return null;
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs>" +
                   "<tab idMso=\"TabMail\"><group id=\"HatirMailGroup\" label=\"HatırMail\">" +
                   "<button id=\"HatirMailSearch\" label=\"HatırMail ile Ara\" size=\"large\" imageMso=\"FindDialog\"" +
                   " screentip=\"Hatırladıklarınızla e-posta bulun\" onAction=\"OnSearch\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        public void OnSearch([MarshalAs(UnmanagedType.IDispatch)] object control)
        {
            try
            {
                if (form == null || form.IsDisposed)
                    form = new SearchForm(new OutlookSource(app), Settings.Load());
                form.Show(); // a window minimized during an index run comes back
                if (form.WindowState == FormWindowState.Minimized) form.WindowState = FormWindowState.Normal;
                form.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("HatırMail açılamadı: " + ex.GetType().Name, "HatırMail", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
            }
        }
    }

    // ---------------------------------------------------------------------------------------------
    // Settings and paths. Everything lives under %LOCALAPPDATA%\HatirMail unless HATIRMAIL_HOME is set.
    // ---------------------------------------------------------------------------------------------
    public class Settings
    {
        public string Home, Python, Worker, Model, Data;
        public int Threads = 4;
        public int MemoryLimitMb = 4096;
        public int SearchIndexBudgetS = 20; // indexing seconds in one search (0 = none); full indexing: İndeks tab
        public int IdleTimeoutS = 300;      // worker gives up if the add-in stops talking
        public int SessionTimeoutS = 7200;  // hard limit for one search (an index run has none: it stops when idle)
        public bool Attachments = true;     // read attachment content (Excel, PDF, Word, pictures with OCR, ...)
        public int KeepAliveMin = 10;       // the model stays loaded so long after the last work while the window is open (0 = no)
        public string ConfigFile { get { return Path.Combine(Home, "settings.json"); } }

        public static Settings Load()
        {
            var s = new Settings();
            s.Home = Environment.GetEnvironmentVariable("HATIRMAIL_HOME");
            if (string.IsNullOrEmpty(s.Home))
                s.Home = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HatirMail");
            s.Python = Path.Combine(s.Home, "runtime", "python", "python.exe");
            s.Worker = Path.Combine(s.Home, "app", "worker", "hm_worker.py");
            s.Model = Path.Combine(s.Home, "model");
            s.Data = Environment.GetEnvironmentVariable("HATIRMAIL_DATA");
            if (string.IsNullOrEmpty(s.Data)) s.Data = Path.Combine(s.Home, "data");
            if (File.Exists(s.ConfigFile))
            {
                try
                {
                    var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(s.ConfigFile, Encoding.UTF8));
                    object v;
                    if (d.TryGetValue("threads", out v)) s.Threads = Convert.ToInt32(v);
                    if (d.TryGetValue("memory_limit_mb", out v)) s.MemoryLimitMb = Convert.ToInt32(v);
                    if (d.TryGetValue("search_index_budget_s", out v)) s.SearchIndexBudgetS = Convert.ToInt32(v);
                    if (d.TryGetValue("idle_timeout_s", out v)) s.IdleTimeoutS = Convert.ToInt32(v);
                    if (d.TryGetValue("session_timeout_s", out v)) s.SessionTimeoutS = Convert.ToInt32(v);
                    if (d.TryGetValue("attachments", out v)) s.Attachments = Convert.ToBoolean(v);
                    if (d.TryGetValue("keep_alive_min", out v)) s.KeepAliveMin = Convert.ToInt32(v);
                }
                catch { } // defaults stay for the values not read
            }
            s.Clamp();
            return s;
        }

        /// <summary>Keep every value in its allowed range (the file can be edited by hand).</summary>
        public void Clamp()
        {
            Threads = Math.Max(1, Math.Min(16, Threads));
            MemoryLimitMb = Math.Max(2048, Math.Min(32768, MemoryLimitMb)); // the model needs ~1.8 GB
            SearchIndexBudgetS = Math.Max(0, Math.Min(600, SearchIndexBudgetS));
            IdleTimeoutS = Math.Max(30, Math.Min(3600, IdleTimeoutS));
            SessionTimeoutS = Math.Max(60, SessionTimeoutS);
            KeepAliveMin = Math.Max(0, Math.Min(120, KeepAliveMin));
        }

        /// <summary>Write settings.json (Ayarlar tab). A temp file first: a failed write leaves the old file.</summary>
        public void Save()
        {
            Clamp();
            var d = new Dictionary<string, object> {
                {"threads", Threads}, {"memory_limit_mb", MemoryLimitMb}, {"search_index_budget_s", SearchIndexBudgetS},
                {"idle_timeout_s", IdleTimeoutS}, {"session_timeout_s", SessionTimeoutS}, {"attachments", Attachments},
                {"keep_alive_min", KeepAliveMin} };
            string tmp = ConfigFile + ".tmp";
            File.WriteAllText(tmp, new JavaScriptSerializer().Serialize(d), new UTF8Encoding(false));
            if (File.Exists(ConfigFile)) File.Replace(tmp, ConfigFile, null);
            else File.Move(tmp, ConfigFile);
        }
    }

    /// <summary>Recent and saved searches (Ara tab): data\searches.json on this PC only. Never sent, never logged.
    /// Uninstall without -KeepData deletes the file; "İndeksi sil" keeps it.</summary>
    public class SearchHistory
    {
        public const int MaxRecent = 20;
        public readonly string FilePath;
        public readonly List<string> Recent = new List<string>();                                       // newest first
        public readonly List<KeyValuePair<string, string>> Saved = new List<KeyValuePair<string, string>>(); // (name, query)
        /// <summary>The file is there but could not be read (locked by an antivirus or backup program): the lists are
        /// empty, and a save would delete the searches in the file. The window reads again before the next change.</summary>
        public bool ReadFailed;

        SearchHistory(string filePath) { FilePath = filePath; }

        public static string PathIn(string dataDir) { return Path.Combine(dataDir, "searches.json"); }

        public SearchHistory Clone()
        {
            var c = new SearchHistory(FilePath);
            c.Recent.AddRange(Recent); c.Saved.AddRange(Saved);
            return c;
        }

        /// <summary>ponytail: a file with bad content (corrupt, or a mistake in a hand edit) gives empty lists, and the
        /// next change writes a new file over it. Items of a wrong type are skipped.</summary>
        public static SearchHistory Load(string dataDir)
        {
            var h = new SearchHistory(PathIn(dataDir));
            string text;
            try { if (!File.Exists(h.FilePath)) return h; text = File.ReadAllText(h.FilePath, Encoding.UTF8); }
            catch { h.ReadFailed = true; return h; }
            try
            {
                var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(text);
                object v;
                if (d.TryGetValue("recent", out v) && v is System.Collections.ArrayList)
                    foreach (var q in ((System.Collections.ArrayList)v).OfType<string>().Reverse()) h.AddRecent(q); // same rules as new searches
                if (d.TryGetValue("saved", out v) && v is System.Collections.ArrayList)
                    foreach (var s in ((System.Collections.ArrayList)v).OfType<Dictionary<string, object>>())
                    {
                        object n, q;
                        if (s.TryGetValue("name", out n) && s.TryGetValue("query", out q)) h.SetSaved(n as string, q as string);
                    }
            }
            catch { h.Recent.Clear(); h.Saved.Clear(); }
            return h;
        }

        /// <summary>Put q first. The same text in another case is one search. At most MaxRecent.
        /// ponytail: ordinal case rules can leave the Turkish ı/I and i/İ unpaired ("faturası" and "FATURASI" can stay two
        /// lines; mono does so); tr-TR rules would break English text ("mail"/"MAIL"). Harmless: one more line in the list.</summary>
        public void AddRecent(string q)
        {
            q = (q ?? "").Trim();
            if (q.Length == 0) return;
            Recent.RemoveAll(x => string.Equals(x, q, StringComparison.OrdinalIgnoreCase));
            Recent.Insert(0, q);
            if (Recent.Count > MaxRecent) Recent.RemoveRange(MaxRecent, Recent.Count - MaxRecent);
        }

        /// <summary>Save query under name. A saved search with the same name (any case) gets the new name and query in its
        /// place: returns true. An empty name or query changes nothing.</summary>
        public bool SetSaved(string name, string query)
        {
            name = (name ?? "").Trim(); query = (query ?? "").Trim();
            if (name.Length == 0 || query.Length == 0) return false;
            int i = Saved.FindIndex(x => string.Equals(x.Key, name, StringComparison.OrdinalIgnoreCase));
            var kv = new KeyValuePair<string, string>(name, query);
            if (i >= 0) Saved[i] = kv; else Saved.Add(kv);
            return i >= 0;
        }

        /// <summary>Write the file (temp file first, as Settings.Save). Makes the data folder if it is not there.</summary>
        public void Save()
        {
            var d = new Dictionary<string, object> {
                {"recent", Recent},
                {"saved", Saved.Select(x => new Dictionary<string, object> { {"name", x.Key}, {"query", x.Value} }).ToList()} };
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
            string tmp = FilePath + ".tmp";
            try
            {
                File.WriteAllText(tmp, new JavaScriptSerializer().Serialize(d), new UTF8Encoding(false));
                if (File.Exists(FilePath)) File.Replace(tmp, FilePath, null);
                else File.Move(tmp, FilePath);
            }
            catch { try { File.Delete(tmp); } catch { } throw; } // no second copy of the search texts stays
        }
    }

    // ---------------------------------------------------------------------------------------------
    // Mail source abstraction: Outlook (late-bound COM) or a JSON file (TestHost, no Outlook).
    // ---------------------------------------------------------------------------------------------
    public class FolderInfo { public string Id, Path, StoreId, StoreName; public int Count; }

    public interface IMailSource
    {
        List<KeyValuePair<string, string>> Stores();                    // (storeId, display name)
        List<FolderInfo> Folders(string storeId);                       // mail folders, recursive
        IEnumerable<object[]> Manifest(FolderInfo f, Func<bool> cancelled); // [eid, lastMod, class, date]
        Dictionary<string, object> Read(string eid, string storeId, string folderId);
        /// <summary>Save attachment number idx (1-based) of a mail as path. Returns false if it is not there.</summary>
        bool SaveAttachment(string eid, string storeId, int idx, string path);
        void Open(string eid, string storeId);
    }

    public class OutlookSource : IMailSource
    {
        readonly dynamic app;
        public OutlookSource(object app) { this.app = app; }

        dynamic Ns { get { return app.Session; } }

        public List<KeyValuePair<string, string>> Stores()
        {
            var r = new List<KeyValuePair<string, string>>();
            dynamic stores = Ns.Stores;
            int n = stores.Count;
            for (int i = 1; i <= n; i++)
            {
                dynamic st = stores[i];
                try { r.Add(new KeyValuePair<string, string>((string)st.StoreID, (string)st.DisplayName)); }
                finally { Release(st); }
            }
            Release(stores);
            return r;
        }

        public List<FolderInfo> Folders(string storeId)
        {
            var r = new List<FolderInfo>();
            dynamic store = Ns.GetStoreFromID(storeId);
            dynamic root = store.GetRootFolder();
            string storeName = store.DisplayName;
            Walk(root, "", storeId, storeName, r, 0);
            Release(root); Release(store);
            return r;
        }

        void Walk(dynamic folder, string prefix, string storeId, string storeName, List<FolderInfo> r, int depth)
        {
            dynamic subs = folder.Folders;
            int n = subs.Count;
            for (int i = 1; i <= n; i++)
            {
                dynamic f = subs[i];
                try
                {
                    string path = prefix + "\\" + (string)f.Name;
                    // 0 = olMailItem: mail folders only (no calendar, contacts, tasks, notes, journal).
                    if ((int)f.DefaultItemType == 0)
                    {
                        dynamic items = f.Items;
                        int count = items.Count;
                        Release(items);
                        r.Add(new FolderInfo { Id = f.EntryID, Path = path, StoreId = storeId, StoreName = storeName, Count = count });
                    }
                    if (depth < 12) Walk(f, path, storeId, storeName, r, depth + 1);
                }
                finally { Release(f); }
            }
            Release(subs);
        }

        public IEnumerable<object[]> Manifest(FolderInfo fi, Func<bool> cancelled)
        {
            dynamic folder = Ns.GetFolderFromID(fi.Id, fi.StoreId);
            dynamic table = folder.GetTable("", 0); // olUserItems; Table reads properties without opening items
            dynamic cols = table.Columns;
            cols.RemoveAll();
            cols.Add("EntryID"); cols.Add("LastModificationTime"); cols.Add("MessageClass"); cols.Add("ReceivedTime");
            Release(cols);
            try
            {
                while (!(bool)table.EndOfTable)
                {
                    if (cancelled()) yield break;
                    dynamic row = table.GetNextRow();
                    object[] v;
                    try
                    {
                        v = new object[] { (string)row[1], Iso(row[2]), (string)row[3], Iso(row[4]) };
                    }
                    finally { Release(row); }
                    yield return v;
                }
            }
            finally { Release(table); Release(folder); }
        }

        public Dictionary<string, object> Read(string eid, string storeId, string folderId)
        {
            dynamic it = Ns.GetItemFromID(eid, storeId);
            try
            {
                string cls = it.MessageClass;
                var d = new Dictionary<string, object>();
                d["eid"] = eid; d["folder"] = folderId; d["cls"] = cls;
                d["mod"] = Iso(it.LastModificationTime);
                d["subject"] = (string)it.Subject ?? "";
                string body = "";
                try { body = it.Body ?? ""; } catch { } // reading Body never changes the item
                d["body"] = body.Length > 200000 ? body.Substring(0, 200000) : body;
                bool isMail = cls != null && cls.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase);
                if (isMail)
                {
                    d["sender"] = Safe(() => (string)it.SenderName);
                    d["sender_addr"] = Safe(() => SmtpOf(it));
                    d["to"] = Safe(() => (string)it.To);
                    d["cc"] = Safe(() => (string)it.CC);
                    d["date"] = FirstDate(Safe(() => (object)it.ReceivedTime), Safe(() => (object)it.SentOn));
                }
                else
                {
                    d["sender"] = Safe(() => (string)it.SenderName);
                    d["sender_addr"] = "";
                    d["to"] = Safe(() => (string)it.To);
                    d["cc"] = "";
                    d["date"] = FirstDate(Safe(() => (object)it.ReceivedTime), Safe(() => (object)it.SentOn),
                                          Safe(() => (object)it.CreationTime));
                }
                // Names for the result list; [index, name, size, type] for attachment reading. Type 1 = file,
                // 5 = attached Outlook item. Size is the MAPI size (a little more than the file), enough for limits.
                var atts = new List<string>();
                var meta = new List<object>();
                dynamic aa = it.Attachments;
                int n = aa.Count;
                for (int i = 1; i <= n; i++)
                {
                    dynamic a = aa[i];
                    try
                    {
                        string name = Safe(() => (string)a.FileName) ?? Safe(() => (string)a.DisplayName) ?? "";
                        if (name.Length > 0) atts.Add(name);
                        meta.Add(new object[] { i, name, Safe(() => (int)a.Size), Safe(() => (int)a.Type) });
                    }
                    finally { Release(a); }
                }
                Release(aa);
                d["attachments"] = atts;
                d["att_meta"] = meta;
                return d;
            }
            finally { Release(it); }
        }

        static string SmtpOf(dynamic it)
        {
            string type = it.SenderEmailType;
            if (type == "EX")
            {
                dynamic s = it.Sender;
                try { dynamic u = s.GetExchangeUser(); if (u != null) { string a = u.PrimarySmtpAddress; Release(u); return a; } }
                finally { Release(s); }
            }
            return it.SenderEmailAddress;
        }

        public bool SaveAttachment(string eid, string storeId, int idx, string path)
        {
            // Attachment.SaveAsFile writes a copy; the mail is not changed and not marked as read.
            dynamic it = Ns.GetItemFromID(eid, storeId);
            try
            {
                dynamic aa = it.Attachments;
                try
                {
                    if (idx < 1 || idx > (int)aa.Count) return false;
                    dynamic a = aa[idx];
                    try { a.SaveAsFile(path); return true; } finally { Release(a); }
                }
                finally { Release(aa); }
            }
            finally { Release(it); }
        }

        public void Open(string eid, string storeId)
        {
            dynamic it = Ns.GetItemFromID(eid, storeId);
            try { it.Display(false); } finally { Release(it); }
        }

        static T Safe<T>(Func<T> f) { try { return f(); } catch { return default(T); } }

        static string FirstDate(params object[] values)
        {
            foreach (var v in values) { string s = Iso(v); if (s.Length > 0) return s; }
            return "";
        }

        static string Iso(object v)
        {
            if (v is DateTime)
            {
                var dt = (DateTime)v;
                if (dt.Year >= 4500 || dt.Year < 1900) return "";
                return dt.ToString("yyyy-MM-dd'T'HH:mm:ss");
            }
            return v == null ? "" : v.ToString();
        }

        static void Release(object o) { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
    }

    /// <summary>JSON test source with the eval/corpus.json schema. Used by TestHost (no Outlook).
    /// A message with "att_paths" (files relative to the corpus file) has readable attachments.</summary>
    public class JsonSource : IMailSource
    {
        readonly Dictionary<string, object> corpus;
        readonly List<Dictionary<string, object>> msgs;
        readonly string dir;
        public readonly List<string> Opened = new List<string>();
        public int Saved;

        public JsonSource(string path)
        {
            var js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            corpus = js.Deserialize<Dictionary<string, object>>(File.ReadAllText(path, Encoding.UTF8));
            msgs = ((System.Collections.ArrayList)corpus["messages"]).Cast<Dictionary<string, object>>().ToList();
            dir = Path.GetDirectoryName(Path.GetFullPath(path));
        }

        List<string> AttPaths(Dictionary<string, object> m)
        {
            return m.ContainsKey("att_paths") ? ((System.Collections.ArrayList)m["att_paths"]).Cast<string>().ToList() : new List<string>();
        }

        Dictionary<string, object> Store { get { return (Dictionary<string, object>)corpus["store"]; } }

        public List<KeyValuePair<string, string>> Stores()
        {
            return new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>((string)Store["id"], (string)Store["name"]) };
        }

        public List<FolderInfo> Folders(string storeId)
        {
            return ((System.Collections.ArrayList)corpus["folders"]).Cast<Dictionary<string, object>>().Select(f =>
                new FolderInfo { Id = (string)f["id"], Path = (string)f["path"], StoreId = storeId, StoreName = (string)Store["name"],
                                 Count = msgs.Count(m => (string)m["folder"] == (string)f["id"]) }).ToList();
        }

        public IEnumerable<object[]> Manifest(FolderInfo f, Func<bool> cancelled)
        {
            foreach (var m in msgs.Where(m => (string)m["folder"] == f.Id))
            {
                if (cancelled()) yield break;
                yield return new object[] { m["id"], m["mod"], m["cls"], m["date"] };
            }
        }

        public Dictionary<string, object> Read(string eid, string storeId, string folderId)
        {
            var m = msgs.First(x => (string)x["id"] == eid);
            Func<string, string> names = k => string.Join("; ", ((System.Collections.ArrayList)m[k]).Cast<System.Collections.ArrayList>().Select(a => (string)a[0]));
            var d = new Dictionary<string, object> {
                {"eid", eid}, {"folder", folderId}, {"cls", m["cls"]}, {"mod", m["mod"]}, {"subject", m["subject"]},
                {"sender", m["from_name"]}, {"sender_addr", m["from_addr"]}, {"to", names("to")}, {"cc", names("cc")},
                {"date", m["date"]}, {"attachments", m["attachments"]}, {"body", m["body"]} };
            var paths = AttPaths(m);
            if (paths.Count > 0)
                d["att_meta"] = paths.Select((p, i) => (object)new object[] {
                    i + 1, Path.GetFileName(p), (int)new FileInfo(Path.Combine(dir, p)).Length, 1 }).ToList();
            return d;
        }

        public bool SaveAttachment(string eid, string storeId, int idx, string path)
        {
            var paths = AttPaths(msgs.First(x => (string)x["id"] == eid));
            if (idx < 1 || idx > paths.Count) return false;
            File.Copy(Path.Combine(dir, paths[idx - 1]), path);
            Saved++;
            return true;
        }

        public void Open(string eid, string storeId) { Opened.Add(eid); }
    }
}

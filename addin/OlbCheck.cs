// OlbCheck: read-only check of the Outlook type library (MSOUTL.OLB).
// It loads the type library with REGKIND_NONE (no registry write, Outlook is not started) and checks that every
// member that HatirMail calls through late binding exists. Usage: HatirMail.OlbCheck.exe <path to MSOUTL.OLB>
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ITypeLib = System.Runtime.InteropServices.ComTypes.ITypeLib;
using ITypeInfo = System.Runtime.InteropServices.ComTypes.ITypeInfo;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;

static class OlbCheck
{
    enum REGKIND { DEFAULT = 0, REGISTER = 1, NONE = 2 }

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern void LoadTypeLibEx(string file, REGKIND kind, out ITypeLib lib);

    static readonly Dictionary<string, string[]> Need = new Dictionary<string, string[]>
    {
        {"_Application", new[] {"Session"}},
        {"_NameSpace", new[] {"Stores", "GetStoreFromID", "GetFolderFromID", "GetItemFromID"}},
        {"_Stores", new[] {"Count", "Item"}},
        {"_Store", new[] {"StoreID", "DisplayName", "GetRootFolder"}},
        {"_Folders", new[] {"Count", "Item"}},
        {"_Items", new[] {"Count"}},
        {"_Table", new[] {"Columns", "EndOfTable", "GetArray"}},
        {"_Columns", new[] {"RemoveAll", "Add"}},
        {"_MailItem", new[] {"MessageClass", "LastModificationTime", "Subject", "Body", "SenderName", "SenderEmailType",
                             "Sender", "SenderEmailAddress", "To", "CC", "ReceivedTime", "SentOn", "Attachments",
                             "Display", "CreationTime", "EntryID", "Parent"}},
        // "Benzerlerini bul": the context of the mail shortcut menu. The PIA has "Selection : IEnumerable" (no "_Selection"
        // base, unlike Folders : _Folders), so the type library name is taken as "Selection"; OlbCheck on 2.19 confirms it.
        {"Selection", new[] {"Count", "Item"}},
        {"MAPIFolder", new[] {"Folders", "Name", "DefaultItemType", "EntryID", "Items", "GetTable", "StoreID"}},
        {"Attachments", new[] {"Count", "Item"}},
        {"Attachment", new[] {"FileName", "DisplayName", "Size", "Type", "SaveAsFile"}},
        {"AddressEntry", new[] {"GetExchangeUser"}},
        {"_ExchangeUser", new[] {"PrimarySmtpAddress"}},
    };

    static int Main(string[] args)
    {
        ITypeLib lib;
        LoadTypeLibEx(args[0], REGKIND.NONE, out lib);
        IntPtr pa;
        lib.GetLibAttr(out pa);
        var la = (TYPELIBATTR)Marshal.PtrToStructure(pa, typeof(TYPELIBATTR));
        lib.ReleaseTLibAttr(pa);
        Console.WriteLine("typelib guid={0} version={1}.{2}", la.guid, la.wMajorVerNum, la.wMinorVerNum);
        var members = new Dictionary<string, HashSet<string>>();
        for (int i = 0; i < lib.GetTypeInfoCount(); i++)
        {
            ITypeInfo ti;
            lib.GetTypeInfo(i, out ti);
            string name, doc, help; int ctx;
            ti.GetDocumentation(-1, out name, out doc, out ctx, out help);
            if (!Need.ContainsKey(name)) continue;
            IntPtr pt;
            ti.GetTypeAttr(out pt);
            var ta = (TYPEATTR)Marshal.PtrToStructure(pt, typeof(TYPEATTR));
            ti.ReleaseTypeAttr(pt);
            var set = new HashSet<string>();
            for (int f = 0; f < ta.cFuncs; f++)
            {
                IntPtr pf;
                ti.GetFuncDesc(f, out pf);
                var fd = (FUNCDESC)Marshal.PtrToStructure(pf, typeof(FUNCDESC));
                string fn, d2, h2; int c2;
                ti.GetDocumentation(fd.memid, out fn, out d2, out c2, out h2);
                set.Add(fn);
                ti.ReleaseFuncDesc(pf);
            }
            members[name] = set;
        }
        int missing = 0;
        foreach (var kv in Need)
        {
            HashSet<string> set;
            if (!members.TryGetValue(kv.Key, out set)) { Console.WriteLine("MISSING TYPE " + kv.Key); missing += kv.Value.Length; continue; }
            var miss = kv.Value.Where(m => !set.Contains(m)).ToList();
            missing += miss.Count;
            Console.WriteLine("{0}: {1}/{2} members found{3}", kv.Key, kv.Value.Length - miss.Count, kv.Value.Length,
                              miss.Count > 0 ? " MISSING: " + string.Join(",", miss) : "");
        }
        Console.WriteLine(missing == 0 ? "PASS: all late-bound members exist" : "FAIL: missing=" + missing);
        return missing == 0 ? 0 : 1;
    }
}

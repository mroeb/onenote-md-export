using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;
using OneNote = Microsoft.Office.Interop.OneNote;

// Exports a local OneNote (desktop) hierarchy to a Markdown folder tree.
//
// No OneDrive / Graph / API permissions required: it drives the installed
// OneNote client through its local COM automation interface.
public static class Program
{
    sealed class PageRef
    {
        public string Id;
        public string Name;
    }

    sealed class SectionRef
    {
        public string Id;
        public string Name;
        public string GroupName;   // may be null
        public List<PageRef> Pages = new List<PageRef>();
    }

    // A page that was actually written, used to build the index.
    sealed class Written
    {
        public string NotebookName;
        public string SectionName;
        public string GroupName;
        public string LinkPath;   // relative to the export root, '/' separated
        public string Title;
    }

    sealed class NotebookRef
    {
        public string Id;
        public string Name;
        public List<SectionRef> Sections = new List<SectionRef>();
    }

    // ---------------- options ----------------
    static string _outDir = "onenote-export";
    static bool _noImages;
    static bool _listOnly;
    static string _notebookFilter;
    static string _sectionFilter;
    static bool _overwrite;
    static bool _skipExisting;
    static bool _dryRun;

    public static int Main(string[] args)
    {
        try
        {
            ParseArgs(args);
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Argument error: " + e.Message);
            PrintUsage();
            return 2;
        }

        if (_listOnly) return ListOnly();

        try
        {
            return Export();
        }
        catch (UnauthorizedAccessException)
        {
            Console.Error.WriteLine("ERROR: access denied. Try running from a normal (non-elevated) shell, and avoid OneNote's own install folder.");
            return 3;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("ERROR: " + e.Message);
            return 1;
        }
    }

    static void ParseArgs(string[] args)
    {
        for (int i = 0; i < args.Length; i++)
        {
            string a = args[i];
            switch (a)
            {
                case "-o":
                case "--out":
                    _outDir = Next(args, ref i, a);
                    break;
                case "--notebook":
                    _notebookFilter = Next(args, ref i, a);
                    break;
                case "--section":
                    _sectionFilter = Next(args, ref i, a);
                    break;
                case "--no-images":
                    _noImages = true;
                    break;
                case "--list":
                    _listOnly = true;
                    break;
                case "--overwrite":
                    _overwrite = true;
                    break;
                case "--skip-existing":
                    _skipExisting = true;
                    break;
                case "-n":
                case "--dry-run":
                    _dryRun = true;
                    break;
                case "-h":
                case "--help":
                    PrintUsage();
                    Environment.Exit(0);
                    break;
                default:
                    if (a.StartsWith("-")) throw new Exception("unknown option " + a);
                    _outDir = a;
                    break;
            }
        }
    }

    static string Next(string[] args, ref int i, string opt)
    {
        if (i + 1 >= args.Length) throw new Exception(opt + " requires a value");
        return args[++i];
    }

    static void PrintUsage()
    {
        Console.WriteLine(@"onenote-md - export local OneNote notebooks to Markdown

Usage:
  onenote-md [output-dir] [options]

Options:
  -o, --out <dir>        Output directory (default: onenote-export)
      --notebook <name>  Only export notebooks whose name contains <name>
      --section <name>   Only export sections whose name contains <name>
      --no-images        Write Markdown without embedding images
      --list             List notebooks/sections/pages, export nothing
      --overwrite        Overwrite existing files (default)
      --skip-existing    Leave existing files untouched
  -n, --dry-run          Show what would be written, write nothing
  -h, --help             This help

Notes:
  * Requires desktop OneNote (Office 2016/2019/365) installed.
  * Reads through the local OneNote client; nothing is sent anywhere.
  * The OneNote window will be driven by the export.");
    }

    // ---------------- COM plumbing ----------------
    [DllImport("ole32.dll")]
    static extern int CoInitializeEx(IntPtr reserved, uint coInit);

    static OneNote.IApplication Connect()
    {
        // STA is required. 0x80010106 == already initialised in another
        // apartment model; in that case just use the existing apartment.
        CoInitializeEx(IntPtr.Zero, 2 /*COINIT_APARTMENTTHREADED*/);

        object raw;
        try
        {
            raw = Activator.CreateInstance(Type.GetTypeFromProgID("OneNote.Application"));
        }
        catch (Exception e)
        {
            throw new Exception(
                "Could not start OneNote automation (" + e.GetBaseException().Message + "). " +
                "Make sure desktop OneNote is installed (not only the Microsoft Store/UWP app).", e);
        }

        OneNote.IApplication app = raw as OneNote.IApplication;
        if (app == null)
            throw new Exception("The OneNote automation object could not be created.");
        return app;
    }

    // ---------------- listing ----------------
    static int ListOnly()
    {
        OneNote.IApplication app = Connect();
        try
        {
            List<NotebookRef> nbs = ReadHierarchy(app);
            int pages = 0;
            foreach (NotebookRef nb in nbs)
            {
                int np = 0;
                foreach (SectionRef s in nb.Sections) np += s.Pages.Count;
                pages += np;
                Console.WriteLine("Notebook: " + nb.Name + "  (" + nb.Sections.Count + " sections, " + np + " pages)");
                foreach (SectionRef s in nb.Sections)
                {
                    string g = s.GroupName == null ? "" : " [group: " + s.GroupName + "]";
                    Console.WriteLine("    Section: " + s.Name + g + "  (" + s.Pages.Count + " pages)");
                    foreach (PageRef p in s.Pages)
                        Console.WriteLine("        - " + p.Name);
                }
            }
            Console.WriteLine();
            Console.WriteLine("Total: " + nbs.Count + " notebooks, " + pages + " pages.");
            return 0;
        }
        finally { Release(app); }
    }

    static void Release(OneNote.IApplication app)
    {
        try { if (app != null) Marshal.FinalReleaseComObject(app); } catch { }
    }

    static List<NotebookRef> ReadHierarchy(OneNote.IApplication app)
    {
        string xml = null;
        app.GetHierarchy("", OneNote.HierarchyScope.hsPages, out xml);
        if (string.IsNullOrEmpty(xml))
            throw new Exception("OneNote returned an empty hierarchy. Is OneNote signed in?");

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(xml);
        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("one", OneNoteConverter.NS);

        List<NotebookRef> result = new List<NotebookRef>();
        foreach (XmlNode nbNode in doc.SelectNodes("/one:Notebooks/one:Notebook", ns))
        {
            NotebookRef nb = new NotebookRef
            {
                Id = TextRunExtractor.Attr(nbNode, "ID"),
                Name = TextRunExtractor.Attr(nbNode, "name")
            };
            CollectSections(nbNode, nb, ns, null);
            result.Add(nb);
        }
        return result;
    }

    static void CollectSections(XmlNode parent, NotebookRef nb, XmlNamespaceManager ns, string groupName)
    {
        foreach (XmlNode child in parent.ChildNodes)
        {
            if (child.NodeType != XmlNodeType.Element) continue;

            if (child.LocalName == "SectionGroup")
            {
                string g = TextRunExtractor.Attr(child, "name");
                CollectSections(child, nb, ns, g ?? groupName);
            }
            else if (child.LocalName == "Section")
            {
                SectionRef sec = new SectionRef
                {
                    Id = TextRunExtractor.Attr(child, "ID"),
                    Name = TextRunExtractor.Attr(child, "name"),
                    GroupName = groupName
                };
                foreach (XmlNode pg in child.SelectNodes("one:Page", ns))
                {
                    sec.Pages.Add(new PageRef
                    {
                        Id = TextRunExtractor.Attr(pg, "ID"),
                        Name = TextRunExtractor.Attr(pg, "name")
                    });
                }
                nb.Sections.Add(sec);
            }
        }
    }

    // ---------------- export ----------------
    static int Export()
    {
        OneNote.IApplication app = Connect();
        int written = 0, failed = 0, skipped = 0;

        try
        {
            List<NotebookRef> nbs = ReadHierarchy(app);
            if (nbs.Count == 0)
            {
                Console.Error.WriteLine("No notebooks found.");
                return 1;
            }

            string root = Path.GetFullPath(_outDir);
            if (!_dryRun) Directory.CreateDirectory(root);

            List<Written> writtenPages = new List<Written>();

            foreach (NotebookRef nb in nbs)
            {
                if (_notebookFilter != null &&
                    nb.Name.IndexOf(_notebookFilter, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                string nbDir = Path.Combine(root, SafeName(nb.Name));
                if (!_dryRun) Directory.CreateDirectory(nbDir);

                foreach (SectionRef sec in nb.Sections)
                {
                    if (_sectionFilter != null &&
                        sec.Name.IndexOf(_sectionFilter, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    if (sec.Pages.Count == 0) continue;

                    string secDir = nbDir;
                    if (sec.GroupName != null)
                        secDir = Path.Combine(Path.Combine(nbDir, SafeName(sec.GroupName)), SafeName(sec.Name));
                    else
                        secDir = Path.Combine(nbDir, SafeName(sec.Name));

                    string assetsDir = _noImages ? null : Path.Combine(secDir, "assets");
                    if (!_dryRun) Directory.CreateDirectory(secDir);

                    // One converter and one assets folder per section. Image names
                    // carry the page's position so pages cannot overwrite each
                    // other's images.
                    int pageOrdinal = 0;
                    OneNoteConverter conv = new OneNoteConverter(assetsDir,
                        MakeAssetNameFor(() => pageOrdinal));

                    // Duplicate page titles are legal, so file names are
                    // de-duplicated per section to avoid losing a page.
                    HashSet<string> usedStems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    int n = 0;
                    foreach (PageRef page in sec.Pages)
                    {
                        pageOrdinal++;
                        string title = string.IsNullOrEmpty(page.Name) ? "Untitled page" : page.Name;

                        string stem = SafeName(title);
                        string unique = stem;
                        int dup = 2;
                        while (!usedStems.Add(unique))
                        {
                            unique = stem + " (" + dup.ToString(CultureInfo.InvariantCulture) + ")";
                            dup++;
                        }

                        string file = Path.Combine(secDir, unique + ".md");

                        if (File.Exists(file) && _skipExisting && !_overwrite) { skipped++; continue; }

                        if (_dryRun)
                        {
                            Console.WriteLine("would write: " + file);
                            n++; written++;
                            continue;
                        }

                        try
                        {
                            string pageXml = null;
                            app.GetPageContent(page.Id, out pageXml,
                                _noImages ? OneNote.PageInfo.piBasic : OneNote.PageInfo.piAll);

                            if (string.IsNullOrEmpty(pageXml))
                            {
                                failed++;
                                continue;
                            }

                            conv.ResetImages();
                            string md = conv.ConvertPage(pageXml, title);
                            md = BuildFrontMatter(page, nb, sec) + md;

                            File.WriteAllText(file, md, new UTF8Encoding(false));
                            writtenPages.Add(new Written
                            {
                                NotebookName = nb.Name,
                                SectionName = sec.Name,
                                GroupName = sec.GroupName,
                                Title = title,
                                LinkPath = Link(
                                    SafeName(nb.Name) + "/"
                                  + (sec.GroupName != null ? SafeName(sec.GroupName) + "/" : "")
                                  + SafeName(sec.Name) + "/" + unique + ".md")
                            });
                            n++; written++;
                            if (n % 10 == 0) Console.WriteLine("  " + nb.Name + " / " + sec.Name + ": " + n + "/" + sec.Pages.Count);
                        }
                        catch (Exception e)
                        {
                            failed++;
                            Console.Error.WriteLine("  ! " + title + ": " + e.GetBaseException().Message);
                        }
                    }
                    Console.WriteLine("  " + secDir.Substring(root.Length).TrimStart('\\') + "  (" + n + " pages)");
                }
            }

            if (!_dryRun) WriteIndex(root, writtenPages);

            Console.WriteLine();
            Console.WriteLine("Done. written=" + written + " failed=" + failed + " skipped=" + skipped);
            Console.WriteLine("Output: " + root);
            return failed == 0 ? 0 : 4;
        }
        finally { Release(app); }
    }

    /// <summary>
    /// Builds a relative Markdown link target from a '/'-separated path.
    /// Each segment is percent-encoded, which is required because OneNote page
    /// titles routinely contain spaces, '(' and ')' - all of which would
    /// otherwise break the Markdown link.
    /// </summary>
    static string Link(string relPath)
    {
        string[] parts = relPath.Split('/');
        for (int i = 0; i < parts.Length; i++)
            parts[i] = EscapeSegment(parts[i]);
        return string.Join("/", parts);
    }

    /// <summary>
    /// Percent-encodes one path segment for use inside a Markdown link.
    /// Uri.EscapeDataString leaves '(' and ')' untouched, but a raw ')' closes
    /// the link target early, so those are encoded explicitly.
    /// </summary>
    static string EscapeSegment(string s)
    {
        return Uri.EscapeDataString(s)
                 .Replace("(", "%28")
                 .Replace(")", "%29")
                 .Replace("!", "%21")
                 .Replace("'", "%27")
                 .Replace("*", "%2A");
    }

    static Func<string, int, string> MakeAssetNameFor(Func<int> pageOrdinal)
    {
        return delegate (string format, int index)
        {
            string ext = (format ?? "png").Trim().ToLowerInvariant();
            if (ext.Length == 0 || ext.Length > 5) ext = "png";
            if (ext == "jpeg") ext = "jpg";
            if (ext == "tif") ext = "tiff";
            return "p" + pageOrdinal().ToString("D2") + "-img" + index.ToString("D3") + "." + ext;
        };
    }

    static string BuildFrontMatter(PageRef page, NotebookRef nb, SectionRef sec)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("---\n");
        sb.Append("title: \"").Append(YamlString(page.Name ?? "")).Append("\"\n");
        sb.Append("notebook: \"").Append(YamlString(nb.Name ?? "")).Append("\"\n");
        sb.Append("section: \"").Append(YamlString(sec.Name ?? "")).Append("\"\n");
        if (sec.GroupName != null)
            sb.Append("group: \"").Append(YamlString(sec.GroupName)).Append("\"\n");
        sb.Append("onenote_id: \"").Append(page.Id ?? "").Append("\"\n");
        sb.Append("---\n\n");
        return sb.ToString();
    }

    static void WriteIndex(string root, List<Written> pages)
    {
        try
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("# OneNote export\n\n");
            sb.Append("_" + pages.Count + " pages._\n");

            string curNotebook = null, curSection = null;
            foreach (Written w in pages)
            {
                if (w.NotebookName != curNotebook)
                {
                    curNotebook = w.NotebookName;
                    curSection = null;
                    sb.Append("\n## ").Append(curNotebook).Append("\n");
                }
                string secKey = (w.GroupName == null ? "" : w.GroupName + "/") + w.SectionName;
                if (secKey != curSection)
                {
                    curSection = secKey;
                    sb.Append("\n### ").Append(w.GroupName == null ? w.SectionName : w.GroupName + " / " + w.SectionName).Append("\n\n");
                }
                sb.Append("- [").Append((w.Title ?? "").Replace("\n", " ")).Append("](")
                  .Append(w.LinkPath).Append(")\n");
            }

            File.WriteAllText(Path.Combine(root, "README.md"), sb.ToString(), new UTF8Encoding(false));
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("Could not write README.md: " + e.Message);
        }
    }

    static string YamlString(string s)
    {
        return (s ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
    }

    /// <summary>Makes a string safe to use as a single path segment.</summary>
    static string SafeName(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "Untitled";
        var sb = new StringBuilder();
        char[] invalid = Path.GetInvalidFileNameChars();
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c < 32 || Array.IndexOf(invalid, c) >= 0) { sb.Append('_'); continue; }
            if (c == ':' || c == '*' || c == '?' || c == '"' || c == '<' || c == '>' || c == '|') { sb.Append('_'); continue; }
            sb.Append(c);
        }
        string r = sb.ToString().Trim().TrimEnd('.');
        if (r.Length == 0) r = "Untitled";
        if (r.Length > 120) r = r.Substring(0, 120).TrimEnd();
        return r;
    }
}

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

// Converts OneNote 2013 page XML into Markdown.
public sealed class OneNoteConverter
{
    public const string NS = "http://schemas.microsoft.com/office/onenote/2013/onenote";

    readonly string _assetsDir;          // where images are written (may be null => skip)
    readonly Func<string, int, string> _assetNamer;
    int _imageCounter;

    // Maps quickStyleIndex -> style name (e.g. "h2", "p", "cite")
    readonly Dictionary<int, string> _styleNames = new Dictionary<int, string>();

    // Objects already rendered during the normal walk. Keyed by node reference
    // so the post-pass can tell "handled" from "missed".
    readonly List<XmlNode> _emitted = new List<XmlNode>();

    // Attachment file names already used in the current section's assets
    // folder, de-duplicated so two attachments of the same name cannot
    // overwrite each other.
    readonly HashSet<string> _usedAttachmentNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    int _attachmentCounter;

    /// <summary>
    /// Clears per-page state. The converter is shared by every page in a
    /// section (they share one assets folder), so this keeps image numbering
    /// starting at 1 for each page and stops the emitted list from growing
    /// across the whole section. Attachment names deliberately survive, since
    /// they must stay unique for the whole section.
    /// </summary>
    public void ResetImages()
    {
        _imageCounter = 0;
        _emitted.Clear();
    }

    public OneNoteConverter(string assetsDir, Func<string, int, string> assetNamer)
    {
        _assetsDir = assetsDir;
        _assetNamer = assetNamer;
    }

    public string ConvertPage(string pageXml, string pageTitle)
    {
        XmlDocument doc = new XmlDocument();
        doc.PreserveWhitespace = false;
        doc.LoadXml(pageXml);

        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("one", NS);

        LoadStyleNames(doc, ns);

        StringBuilder md = new StringBuilder();
        md.Append("# ").Append(EscapeTitle(pageTitle)).Append("\n\n");

        // Page title as authored inside the page (may differ from hierarchy name).
        XmlNode titleNode = doc.SelectSingleNode("/one:Page/one:Title", ns);
        if (titleNode != null)
        {
            string inner = PlainText(titleNode);
            if (!string.IsNullOrEmpty(inner) && inner.Trim() != pageTitle.Trim())
                md.Append("**").Append(inner.Trim()).Append("**\n\n");
        }

        foreach (XmlNode outline in doc.SelectNodes("/one:Page/one:Outline", ns))
            WriteOutline(outline, md, ns, 0);

        // Images and attachments do not always sit inside an outline: OneNote
        // also places them directly under <one:Page> (pinned/floating objects)
        // and occasionally elsewhere. Anything the walk above did not emit is
        // appended here in document order, so nothing is silently dropped.
        AppendUnemitted(doc, md, ns);

        return md.ToString();
    }

    void AppendUnemitted(XmlDocument doc, StringBuilder md, XmlNamespaceManager ns)
    {
        // A single XPath returns both kinds in document order, which is both
        // cheaper and simpler than collecting two lists and sorting afterwards.
        XmlNodeList candidates = doc.SelectNodes("//*[local-name()='Image' or @preferredName]", ns);
        if (candidates == null || candidates.Count == 0) return;

        foreach (XmlNode node in candidates)
        {
            if (AlreadyEmitted(node)) continue;

            if (node.LocalName == "Image")
            {
                // Callback-only or empty payloads are not recoverable; an empty
                // <one:Data/> is a print artefact rather than a real picture.
                XmlNode data = node.SelectSingleNode("one:Data", ns);
                if (data == null || data.InnerText.Trim().Length == 0) continue;
            }

            string rendered = node.LocalName == "Image"
                ? RenderImage(node, ns)
                : RenderAttachment(node);

            if (rendered.Length == 0) continue;      // renderers record it either way
            if (!EndsWithBlankLine(md)) md.Append("\n");
            md.Append(rendered).Append("\n\n");
        }
    }

    static bool EndsWithBlankLine(StringBuilder md)
    {
        if (md.Length < 2) return false;
        return md[md.Length - 1] == '\n' && md[md.Length - 2] == '\n';
    }

    bool AlreadyEmitted(XmlNode n)
    {
        for (int i = 0; i < _emitted.Count; i++)
            if (ReferenceEquals(_emitted[i], n)) return true;
        return false;
    }

    void LoadStyleNames(XmlDocument doc, XmlNamespaceManager ns)
    {
        foreach (XmlNode q in doc.SelectNodes("/one:Page/one:QuickStyleDef", ns))
        {
            string idx = TextRunExtractor.Attr(q, "index");
            string nm = TextRunExtractor.Attr(q, "name");
            int i;
            if (idx != null && nm != null && int.TryParse(idx, out i))
                _styleNames[i] = nm;
        }
    }

    // <one:Outline> == one bullet level. Depending on the schema revision the
    // actual <one:OE> paragraphs hang directly off the Outline or off an
    // intermediate <one:OEChildren>, so both shapes are handled here.
    void WriteOutline(XmlNode outline, StringBuilder md, XmlNamespaceManager ns, int depth)
    {
        foreach (XmlNode child in outline.ChildNodes)
        {
            if (child.NodeType != XmlNodeType.Element) continue;

            switch (child.LocalName)
            {
                case "OE":
                    WriteOE(child, md, ns, depth);
                    break;

                case "OEChildren":
                    foreach (XmlNode oe in child.SelectNodes("one:OE", ns))
                        WriteOE(oe, md, ns, depth);
                    foreach (XmlNode nested in child.SelectNodes("one:Outline", ns))
                        WriteOutline(nested, md, ns, depth);
                    break;

                case "Outline":
                    WriteOutline(child, md, ns, depth);
                    break;
            }
        }
    }

    void WriteOE(XmlNode oe, StringBuilder md, XmlNamespaceManager ns, int depth)
    {
        string quickStyle = TextRunExtractor.Attr(oe, "quickStyleIndex");
        string styleName = null;
        int qi;
        if (quickStyle != null && int.TryParse(quickStyle, out qi) && _styleNames.ContainsKey(qi))
            styleName = _styleNames[qi];

        int level = HeadingLevel(styleName);
        bool isBullet = oe.SelectSingleNode("one:List", ns) != null;
        string pad = new string(' ', depth * 2);

        if (level > 0)
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
            {
                md.Append(pad);
                for (int h = 0; h < level; h++) md.Append("#");
                md.Append(" ").Append(text.Trim()).Append("\n\n");
            }
        }
        else if (isBullet)
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
                md.Append(pad).Append("- ").Append(text.Trim()).Append("\n");
        }
        else
        {
            string text = InlineContent(oe, ns);
            if (text.Trim().Length > 0)
                md.Append(text.TrimEnd()).Append("\n\n");
        }

        // Tables rendered by OneNote arrive as an HTML blob.
        XmlNode table = oe.SelectSingleNode("one:OEChildren/one:T[@tableHTML] | one:T[@tableHTML]", ns);
        if (table != null)
        {
            string html = TextRunExtractor.Attr(table, "tableHTML");
            if (!string.IsNullOrEmpty(html))
            {
                md.Append(ConvertTableHtml(html)).Append("\n\n");
            }
        }

        XmlNode kids = oe.SelectSingleNode("one:OEChildren", ns);
        if (kids != null)
        {
            foreach (XmlNode childOE in kids.SelectNodes("one:OE", ns))
                WriteOE(childOE, md, ns, depth + 1);
            foreach (XmlNode childOutline in kids.SelectNodes("one:Outline", ns))
                WriteOutline(childOutline, md, ns, depth + 1);
        }
    }

    static int HeadingLevel(string styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return 0;
        if (styleName == "PageTitle") return 0;         // title handled separately
        if (styleName.Length == 2 && styleName[0] == 'h' && styleName[1] >= '1' && styleName[1] <= '6')
            return styleName[1] - '0';
        return 0;
    }

    // Renders the inline content of an <one:OE> as a single line of Markdown.
    string InlineContent(XmlNode oe, XmlNamespaceManager ns)
    {
        StringBuilder sb = new StringBuilder();

        foreach (XmlNode node in oe.ChildNodes)
        {
            if (node.NodeType != XmlNodeType.Element) continue;

            switch (node.LocalName)
            {
                case "T":
                    {
                        string raw = InnerTextOf(node);
                        if (LooksLikeHtml(raw))
                            sb.Append(ConvertInlineHtml(raw));
                        else
                        {
                            var tr = new TextRunExtractor();
                            tr.Extract(node);
                            sb.Append(tr.Out.ToString());
                        }
                        break;
                    }
                case "Image":
                    sb.Append(RenderImage(node, ns));
                    break;
                case "InsertedFile":
                case "MediaFile":
                case "File":
                    sb.Append(RenderAttachment(node));
                    break;
                case "OEChildren":
                    break; // handled by caller
            }
        }

        return sb.ToString();
    }

    string InnerTextOf(XmlNode t)
    {
        if (t == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (XmlNode c in t.ChildNodes)
        {
            if (c.NodeType == XmlNodeType.CDATA || c.NodeType == XmlNodeType.Text)
                sb.Append(c.Value ?? "");
            else if (c.LocalName == "r")
            {
                foreach (XmlNode rc in c.ChildNodes)
                    if (rc.NodeType == XmlNodeType.CDATA || rc.NodeType == XmlNodeType.Text)
                        sb.Append(rc.Value ?? "");
            }
        }
        // CDATA holds text verbatim, so HTML entities inside it are still encoded.
        return DecodeEntities(sb.ToString());
    }

    /// <summary>
    /// Resolves the HTML/XML entities OneNote leaves inside CDATA text. Entities
    /// are resolved left to right, so "&amp;lt;" becomes "&lt;" rather than
    /// collapsing all the way down to "<".
    /// </summary>
    public static string DecodeEntities(string s)
    {
        if (string.IsNullOrEmpty(s) || s.IndexOf('&') < 0) return s;

        StringBuilder sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '&') { sb.Append(s[i]); continue; }

            int semi = s.IndexOf(';', i + 1);
            if (semi < 0 || semi - i > 10) { sb.Append('&'); continue; }

            string ent = s.Substring(i + 1, semi - i - 1);
            string rep = null;

            if (ent.Length > 1 && ent[0] == '#')
            {
                int cp;
                bool ok = ent[1] == 'x' || ent[1] == 'X'
                    ? int.TryParse(ent.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out cp)
                    : int.TryParse(ent.Substring(1), out cp);
                if (ok && cp > 0 && cp <= 0x10FFFF)
                {
                    // Soft line breaks inside a paragraph become spaces.
                    if (cp == 0x0A || cp == 0x0D) sb.Append(' ');
                    else sb.Append(char.ConvertFromUtf32(cp));
                    i = semi;
                    continue;
                }
            }
            else
            {
                switch (ent)
                {
                    case "amp": rep = "&"; break;
                    case "lt": rep = "<"; break;
                    case "gt": rep = ">"; break;
                    case "quot": rep = "\""; break;
                    case "apos": rep = "'"; break;
                    case "nbsp": rep = " "; break;
                    case "hellip": rep = "\u2026"; break;
                    case "ndash": rep = "\u2013"; break;
                    case "mdash": rep = "\u2014"; break;
                    case "bull": rep = "\u2022"; break;
                    case "auml": rep = "\u00e4"; break;
                    case "ouml": rep = "\u00f6"; break;
                    case "uuml": rep = "\u00fc"; break;
                    case "Auml": rep = "\u00c4"; break;
                    case "Ouml": rep = "\u00d6"; break;
                    case "Uuml": rep = "\u00dc"; break;
                    case "szlig": rep = "\u00df"; break;
                    case "euro": rep = "\u20ac"; break;
                    case "copy": rep = "\u00a9"; break;
                    case "reg": rep = "\u00ae"; break;
                    case "deg": rep = "\u00b0"; break;
                    case "plusmn": rep = "\u00b1"; break;
                    case "times": rep = "\u00d7"; break;
                    case "middot": rep = "\u00b7"; break;
                    case "laquo": rep = "\u00ab"; break;
                    case "raquo": rep = "\u00bb"; break;
                    case "ldquo": rep = "\u201c"; break;
                    case "rdquo": rep = "\u201d"; break;
                    case "lsquo": rep = "\u2018"; break;
                    case "rsquo": rep = "\u2019"; break;
                    case "eacute": rep = "\u00e9"; break;
                    case "egrave": rep = "\u00e8"; break;
                    case "agrave": rep = "\u00e0"; break;
                }
            }

            if (rep != null) { sb.Append(rep); i = semi; }
            else sb.Append('&');
        }
        return sb.ToString();
    }

    // OneNote often stores a whole <one:T> as a literal HTML fragment that
    // mixes plain text and markup, so the test is for any inline tag anywhere
    // in the string rather than one at the very start.
    static readonly string[] HtmlTags =
    {
        "<span", "<b>", "<b ", "<i>", "<i ", "<em", "<font", "<a ", "<a>",
        "<br", "<code", "<sub", "<sup", "<u>", "<strong", "<s>", "<strike", "<img"
    };

    static bool LooksLikeHtml(string s)
    {
        if (string.IsNullOrEmpty(s)) return false;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '<') continue;
            for (int t = 0; t < HtmlTags.Length; t++)
            {
                string tag = HtmlTags[t];
                if (i + tag.Length <= s.Length &&
                    string.Compare(s, i, tag, 0, tag.Length, StringComparison.OrdinalIgnoreCase) == 0)
                    return true;
            }
        }
        return false;
    }

    string RenderImage(XmlNode img, XmlNamespaceManager ns)
    {
        string alt = TextRunExtractor.Attr(img, "alt");
        string format = TextRunExtractor.Attr(img, "format");
        if (string.IsNullOrEmpty(format)) format = "png";

        XmlNode data = img.SelectSingleNode("one:Data", ns);
        if (data == null) return "";

        string b64 = data.InnerText;
        if (string.IsNullOrEmpty(b64)) return "";

        if (_assetsDir == null) return "";   // images suppressed

        try
        {
            byte[] bytes = Convert.FromBase64String(b64);
            _imageCounter++;
            string fname = _assetNamer(format, _imageCounter);
            Directory.CreateDirectory(_assetsDir);
            File.WriteAllBytes(Path.Combine(_assetsDir, fname), bytes);
            _emitted.Add(img);
            return "![" + EscapeAlt(ShortAlt(alt)) + "](assets/" + fname + ")";
        }
        catch (Exception)
        {
            // Mark as handled so the post-pass does not retry a broken payload.
            _emitted.Add(img);
            return "";
        }
    }

    /// <summary>
    /// OneNote stores the full OCR transcript in the image's alt attribute,
    /// which would otherwise become an unreadable multi-kilobyte alt text.
    /// </summary>
    static string ShortAlt(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "image";
        string t = Collapse(s);
        const int max = 120;
        if (t.Length <= max) return t;
        int cut = t.LastIndexOf(' ', max);
        if (cut < 40) cut = max;
        return t.Substring(0, cut).TrimEnd() + "\u2026";
    }

    static string Collapse(string s)
    {
        return s.Replace("\r", " ").Replace("\n", " ").Replace("\u00A0", " ").Trim();
    }

    static string EscapeAlt(string s)
    {
        return s.Replace("[", "\\[").Replace("]", "\\]");
    }

    static string EscapeTitle(string s)
    {
        if (string.IsNullOrEmpty(s)) return "Untitled";
        return s.Replace("\r", " ").Replace("\n", " ").Trim();
    }

    static string PlainText(XmlNode n)
    {
        if (n == null) return "";
        StringBuilder sb = new StringBuilder();
        foreach (XmlNode c in n.ChildNodes)
        {
            if (c.NodeType == XmlNodeType.CDATA || c.NodeType == XmlNodeType.Text) sb.Append(c.Value ?? "");
            else if (c.LocalName == "r" || c.LocalName == "T")
                sb.Append(PlainText(c));
        }
        return sb.ToString();
    }

    /// <summary>
    /// Renders an <one:InsertedFile> / <one:MediaFile> attachment. OneNote
    /// keeps the bytes in its own cache and points at them with pathCache;
    /// preferredName is the name the file had when it was attached, which is
    /// what a reader expects to see.
    /// </summary>
    string RenderAttachment(XmlNode node)
    {
        _emitted.Add(node);

        if (_assetsDir == null) return "";

        string preferred = TextRunExtractor.Attr(node, "preferredName");
        string cachePath = TextRunExtractor.Attr(node, "pathCache");

        if (string.IsNullOrEmpty(preferred) && string.IsNullOrEmpty(cachePath))
            return "";

        _attachmentCounter++;
        string ext = SafeExtension(preferred);
        string baseName = SafeFileNameWithoutExtension(
            string.IsNullOrEmpty(preferred) ? Path.GetFileName(cachePath) : preferred);

        if (baseName.Length == 0)
            baseName = "attachment-" + _attachmentCounter.ToString("D3");

        // Guarantee a unique name within the section.
        string candidate = baseName + ext;
        int dup = 2;
        while (!_usedAttachmentNames.Add(candidate))
        {
            candidate = baseName + " (" + dup.ToString(CultureInfo.InvariantCulture) + ")" + ext;
            dup++;
        }

        string label = baseName + ext;
        string sizeText = "";

        if (!string.IsNullOrEmpty(cachePath) && File.Exists(cachePath))
        {
            try
            {
                string dest = Path.Combine(_assetsDir, candidate);
                Directory.CreateDirectory(_assetsDir);
                File.Copy(cachePath, dest, true);
                sizeText = " (" + HumanSize(new FileInfo(dest).Length) + ")";
            }
            catch (Exception)
            {
                return "[" + EscapeLabel(label) + "] (attachment could not be copied)";
            }
        }
        else
        {
            // Not synced to this machine yet, so the bytes are unavailable.
            return "[" + EscapeLabel(label) + "] (not available locally)";
        }

        return "[" + EscapeLabel(label + sizeText) + "](assets/" + Uri.EscapeDataString(candidate)
                 .Replace("(", "%28").Replace(")", "%29") + ")";
    }

    static string EscapeLabel(string s)
    {
        return s.Replace("[", "\\[").Replace("]", "\\]");
    }

    static string SafeExtension(string name)
    {
        if (string.IsNullOrEmpty(name)) return "";
        string ext = Path.GetExtension(name);
        if (string.IsNullOrEmpty(ext) || ext.Length > 12) return "";
        foreach (char c in ext)
            if (!(char.IsLetterOrDigit(c) || c == '.')) return "";
        return ext.ToLowerInvariant();
    }

    static string SafeFileNameWithoutExtension(string name)
    {
        if (string.IsNullOrEmpty(name)) return "";
        string stem = Path.GetFileNameWithoutExtension(name);
        if (string.IsNullOrEmpty(stem)) stem = name;
        var sb = new StringBuilder(stem.Length);
        for (int i = 0; i < stem.Length; i++)
        {
            char c = stem[i];
            bool bad = c < 32 || c == '/' || c == '\\' || c == ':' || c == '*' || c == '?'
                       || c == '"' || c == '<' || c == '>' || c == '|' || c == '?';
            sb.Append(bad ? '_' : c);
        }
        string r = sb.ToString().Trim().TrimEnd('.');
        if (r.Length > 100) r = r.Substring(0, 100).TrimEnd();
        return r;
    }

    static string HumanSize(long bytes)
    {
        if (bytes < 1024) return bytes + " B";
        double kb = bytes / 1024.0;
        if (kb < 1024) return kb.ToString("0.#", CultureInfo.InvariantCulture) + " KB";
        double mb = kb / 1024.0;
        if (mb < 1024) return mb.ToString("0.#", CultureInfo.InvariantCulture) + " MB";
        return (mb / 1024.0).ToString("0.##", CultureInfo.InvariantCulture) + " GB";
    }

    // OneNote stores some formatted runs as literal HTML inside <one:T>. Those
    // fragments are frequently not well-formed XML (unclosed tags, stray
    // entities), so this is a tolerant scanner rather than a DOM parse: it
    // never throws and degrades to plain text instead of losing content.
    string ConvertInlineHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";

        StringBuilder sb = new StringBuilder(html.Length);

        // Marks that were opened and still need to be closed at the end.
        List<string> open = new List<string>();
        List<string> links = new List<string>();   // hrefs, aligned with <a>
        string openSpan = null;                    // currently active <span> mark

        int i = 0;
        while (i < html.Length)
        {
            char c = html[i];
            if (c != '<') { sb.Append(c); i++; continue; }

            int gt = html.IndexOf('>', i + 1);
            if (gt < 0) { sb.Append(html.Substring(i)); break; }

            string tag = html.Substring(i + 1, gt - i - 1);
            i = gt + 1;

            if (tag.Length == 0) continue;
            if (tag[0] == '!') continue;                       // comment / doctype

            bool closing = tag[0] == '/';
            if (closing) tag = tag.Substring(1);

            int sp = tag.IndexOfAny(new char[] { ' ', '\t', '\r', '\n', '/' });
            string name = (sp < 0 ? tag : tag.Substring(0, sp)).ToLowerInvariant();
            string attrs = sp < 0 ? "" : tag.Substring(sp);

            switch (name)
            {
                case "br":
                    sb.Append("  \n");
                    break;

                case "b":
                case "strong":
                    ApplyMark(sb, open, "**", closing);
                    break;

                case "i":
                case "em":
                    ApplyMark(sb, open, "*", closing);
                    break;

                case "u":
                case "ins":
                    ApplyMark(sb, open, "<u>", closing);
                    break;

                case "s":
                case "strike":
                case "del":
                    ApplyMark(sb, open, "~~", closing);
                    break;

                case "code":
                    ApplyMark(sb, open, "`", closing);
                    break;

                case "span":
                    {
                        // OneNote emits one <span> per formatting run and often
                        // omits </span>, so treat it as a toggle: the previous run
                        // is always closed before the next one begins.
                        if (openSpan != null) { sb.Append(openSpan); openSpan = null; }
                        if (!closing)
                        {
                            string st = (AttrValue(attrs, "style") ?? "").ToLowerInvariant();
                            if (st.Contains("font-weight:bold") || st.Contains("font-weight:700")) { openSpan = "**"; sb.Append("**"); }
                            else if (st.Contains("font-style:italic")) { openSpan = "*"; sb.Append("*"); }
                            else if (st.Contains("text-decoration:underline")) { openSpan = "<u>"; sb.Append("<u>"); }
                        }
                        break;
                    }

                case "a":
                    {
                        if (closing)
                        {
                            if (links.Count > 0)
                            {
                                string href = links[links.Count - 1];
                                links.RemoveAt(links.Count - 1);
                                if (!string.IsNullOrEmpty(href)) sb.Append(" (").Append(href).Append(")");
                            }
                        }
                        else
                        {
                            links.Add(AttrValue(attrs, "href"));
                        }
                        break;
                    }

                case "img":
                    {
                        string src = AttrValue(attrs, "src");
                        string alt = AttrValue(attrs, "alt");
                        if (!string.IsNullOrEmpty(src))
                            sb.Append("![").Append(alt ?? "image").Append("](").Append(src).Append(")");
                        break;
                    }

                default:
                    break;   // unknown tag: drop the tag, keep the text
            }
        }

        if (openSpan != null) { sb.Append(openSpan); openSpan = null; }
        for (int k = open.Count - 1; k >= 0; k--) sb.Append(open[k]);
        return sb.ToString();
    }

    static void ApplyMark(StringBuilder sb, List<string> open, string mark, bool closing)
    {
        if (closing)
        {
            for (int k = open.Count - 1; k >= 0; k--)
            {
                if (open[k] == mark)
                {
                    sb.Append(mark);
                    open.RemoveRange(k, open.Count - k);
                    return;
                }
            }
        }
        else
        {
            open.Add(mark);
            sb.Append(mark);
        }
    }

    /// <summary>Pulls a value out of a raw HTML attribute list.</summary>
    static string AttrValue(string attrs, string name)
    {
        if (string.IsNullOrEmpty(attrs)) return null;
        int at = 0;
        while ((at = attrs.IndexOf(name, at, StringComparison.OrdinalIgnoreCase)) >= 0)
        {
            // must be preceded by whitespace and followed by '='
            if (at == 0 || char.IsWhiteSpace(attrs[at - 1]))
            {
                int eq = at + name.Length;
                while (eq < attrs.Length && char.IsWhiteSpace(attrs[eq])) eq++;
                if (eq < attrs.Length && attrs[eq] == '=')
                {
                    eq++;
                    while (eq < attrs.Length && char.IsWhiteSpace(attrs[eq])) eq++;
                    if (eq < attrs.Length && (attrs[eq] == '"' || attrs[eq] == '\''))
                    {
                        char q = attrs[eq];
                        int end = attrs.IndexOf(q, eq + 1);
                        if (end > eq) return DecodeEntities(attrs.Substring(eq + 1, end - eq - 1));
                    }
                    else
                    {
                        int end = eq;
                        while (end < attrs.Length && !char.IsWhiteSpace(attrs[end]) && attrs[end] != '>') end++;
                        return DecodeEntities(attrs.Substring(eq, end - eq));
                    }
                }
            }
            at += name.Length;
        }
        return null;
    }

    // Converts a OneNote tableHTML blob into a GitHub-flavoured Markdown table.
    string ConvertTableHtml(string html)
    {
        try
        {
            XmlDocument d = new XmlDocument();
            d.LoadXml("<root>" + html + "</root>");

            List<List<string>> rows = new List<List<string>>();
            XmlNodeList trs = d.GetElementsByTagName("tr");
            foreach (XmlNode tr in trs)
            {
                List<string> cells = new List<string>();
                foreach (XmlNode cell in tr.ChildNodes)
                {
                    if (cell.NodeType != XmlNodeType.Element) continue;
                    if (cell.LocalName != "td" && cell.LocalName != "th") continue;
                    // InnerXml keeps inline markup, which the scanner re-reads so
                    // that bold/italic inside cells survives.
                    cells.Add(CleanCell(ConvertInlineHtml(cell.InnerXml)));
                }
                if (cells.Count > 0) rows.Add(cells);
            }

            if (rows.Count == 0) return "";

            int cols = 0;
            foreach (List<string> r in rows) cols = Math.Max(cols, r.Count);
            if (cols == 0) return "";

            StringBuilder sb = new StringBuilder();
            List<string> header = rows[0];
            sb.Append("| ");
            for (int i = 0; i < cols; i++) sb.Append(EscapeCell(i < header.Count ? header[i] : "")).Append(" | ");
            sb.Append("\n| ");
            for (int i = 0; i < cols; i++) sb.Append(" --- | ");
            sb.Append("\n");

            for (int r = 1; r < rows.Count; r++)
            {
                List<string> row = rows[r];
                sb.Append("| ");
                for (int i = 0; i < cols; i++) sb.Append(EscapeCell(i < row.Count ? row[i] : "")).Append(" | ");
                sb.Append("\n");
            }
            return sb.ToString().TrimEnd('\n');
        }
        catch (Exception)
        {
            return "";
        }
    }

    static string CleanCell(string s)
    {
        return s.Replace("\r", " ").Replace("\n", " ").Replace("\u00A0", " ").Trim();
    }

    static string EscapeCell(string s)
    {
        return (s ?? "").Replace("|", "\\|");
    }
}

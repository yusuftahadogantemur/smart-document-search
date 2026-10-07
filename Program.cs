using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using UglyToad.PdfPig;
using W = DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;

namespace AkilliDokumanArama;

static class Program
{
    [STAThread]
    static void Main()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

record Hit(string Type, string FileName, string Folder, string Snippet);

static class Ocr
{
    static Tesseract.TesseractEngine? engine;
    static bool initTried;
    public static string Error = "";

    public static bool Available
    {
        get { Init(); return engine != null; }
    }

    static void Init()
    {
        if (initTried) return;
        initTried = true;
        try
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "tessdata");
            var langs = new List<string>();
            foreach (string l in new[] { "tur", "eng" })
                if (File.Exists(Path.Combine(dir, l + ".traineddata"))) langs.Add(l);

            if (langs.Count == 0)
            {
                Error = "No tur.traineddata / eng.traineddata found in the tessdata folder.";
                return;
            }
            engine = new Tesseract.TesseractEngine(dir, string.Join("+", langs), Tesseract.EngineMode.Default);
        }
        catch (Exception ex)
        {
            Error = ex.Message;
        }
    }

    public static string? Read(byte[] imageBytes)
    {
        if (!Available) return null;
        try
        {
            using var pix = Tesseract.Pix.LoadFromMemory(imageBytes);
            using var page = engine!.Process(pix);
            return page.GetText();
        }
        catch
        {
            return null;
        }
    }
}

static class SearchEngine
{
    const string OcrMark = "\u0001";
    const long MaxFileBytes = 64L * 1024 * 1024;

    static readonly HashSet<string> TextExts = new()
    {
        ".txt", ".md", ".csv", ".tsv", ".log", ".json", ".xml", ".html", ".htm", ".ini", ".cfg",
        ".yaml", ".yml", ".sql", ".cs", ".py", ".js", ".ts", ".css", ".java", ".c", ".cpp", ".h",
        ".rtf", ".srt", ".tex",
    };
    static readonly HashSet<string> ImageExts = new()
    {
        ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif",
    };
    static readonly HashSet<string> LegacyExts = new()
    {
        ".doc", ".xls", ".ppt", ".msg", ".pub", ".wps",
    };

    public static readonly HashSet<string> SkipDirs = new(StringComparer.OrdinalIgnoreCase)
    {
        "Windows", "Program Files", "Program Files (x86)", "ProgramData", "$Recycle.Bin",
        "System Volume Information", "AppData", "node_modules", ".git", "WinSxS", "$WINDOWS.~BT",
    };

    public static string TypeOf(string ext)
    {
        if (ext is ".docx" or ".doc" or ".odt") return "Word";
        if (ext is ".xlsx" or ".xlsm" or ".xls" or ".ods") return "Excel";
        if (ext is ".pptx" or ".ppt" or ".odp") return "PowerPoint";
        if (ext == ".pdf") return "PDF";
        if (ImageExts.Contains(ext)) return "Image";
        if (TextExts.Contains(ext)) return "Text";
        return ext.Length > 1 ? ext.TrimStart('.').ToUpperInvariant() : "Other";
    }

    public static string Fold(string s)
    {
        var sb = new StringBuilder(s.Length);
        foreach (char c in s)
        {
            sb.Append(c switch
            {
                'İ' or 'I' or 'ı' => 'i',
                'ç' or 'Ç' => 'c',
                'ğ' or 'Ğ' => 'g',
                'ö' or 'Ö' => 'o',
                'ş' or 'Ş' => 's',
                'ü' or 'Ü' => 'u',
                _ => char.ToLowerInvariant(c),
            });
        }
        return sb.ToString();
    }

    public static (string? Snippet, bool Failed) SearchFile(
        string path, string root, string foldedQuery, bool useOcr, CancellationToken token)
    {
        string? snippet = SearchFileCore(path, root, foldedQuery, useOcr, token, out bool failed);
        return (snippet, failed);
    }

    static string? SearchFileCore(string path, string root, string foldedQuery, bool useOcr,
        CancellationToken token, out bool failed)
    {
        failed = false;
        string ext = Path.GetExtension(path).ToLowerInvariant();

        if (Fold(Path.GetFileNameWithoutExtension(path)).Contains(foldedQuery))
            return "(found in file name)";

        string relDir = Path.GetDirectoryName(Path.GetRelativePath(root, path)) ?? "";
        if (relDir.Length > 0 && Fold(relDir).Contains(foldedQuery))
            return "(found in folder name: " + relDir + ")";

        try
        {
            var attrs = File.GetAttributes(path);
            if ((attrs & FileAttributes.Offline) != 0 ||
                ((int)attrs & 0x400000) != 0 ||
                ((int)attrs & 0x40000) != 0)
                return null;
        }
        catch { failed = true; return null; }

        try
        {
            foreach (string chunk in Extract(path, ext, useOcr))
            {
                token.ThrowIfCancellationRequested();
                bool fromOcr = chunk.Length > 0 && chunk[0] == OcrMark[0];
                string folded = Fold(chunk);
                int i = folded.IndexOf(foldedQuery, StringComparison.Ordinal);
                if (i >= 0)
                    return (fromOcr ? "[OCR] " : "") + MakeSnippet(chunk, i, foldedQuery.Length);
            }
        }
        catch
        {
            failed = true;
        }
        return null;
    }

    static string MakeSnippet(string text, int index, int length)
    {
        int start = Math.Max(0, index - 60);
        int end = Math.Min(text.Length, index + length + 60);
        string part = new string(text[start..end].Select(c => char.IsControl(c) ? ' ' : c).ToArray());
        return string.Join(" ", part.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    }

    static IEnumerable<string> Extract(string path, string ext, bool useOcr)
    {
        if (TextExts.Contains(ext)) return ReadText(path);
        if (ImageExts.Contains(ext)) return ReadImage(path, useOcr);
        return ext switch
        {
            ".docx" => ReadDocx(path),
            ".xlsx" or ".xlsm" => ReadXlsx(path),
            ".pptx" => ReadPptx(path),
            ".odt" or ".ods" or ".odp" => ReadOpenDocument(path),
            ".pdf" => ReadPdf(path, useOcr),
            _ => ReadRaw(path, ext),
        };
    }

    static FileStream OpenShared(string path) =>
        new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

    static byte[] ReadAllBytesShared(string path)
    {
        using var fs = OpenShared(path);
        var buf = new byte[fs.Length];
        fs.ReadExactly(buf);
        return buf;
    }

    static string DecodeText(byte[] bytes)
    {
        try { return new UTF8Encoding(false, true).GetString(bytes); }
        catch { return Encoding.GetEncoding(1254).GetString(bytes); }
    }

    static IEnumerable<string> ReadText(string path)
    {
        if (new FileInfo(path).Length > MaxFileBytes) yield break;
        yield return DecodeText(ReadAllBytesShared(path));
    }

    static IEnumerable<string> ReadImage(string path, bool useOcr)
    {
        if (!useOcr) yield break;
        if (new FileInfo(path).Length > MaxFileBytes) yield break;
        string? text = Ocr.Read(ReadAllBytesShared(path));
        if (!string.IsNullOrWhiteSpace(text)) yield return OcrMark + text;
    }

    static IEnumerable<string> ReadDocx(string path)
    {
        using var fs = OpenShared(path);
        using var doc = WordprocessingDocument.Open(fs, false);
        var body = doc.MainDocumentPart?.Document?.Body;
        if (body == null) yield break;
        foreach (var p in body.Descendants<W.Paragraph>())
            yield return p.InnerText;
    }

    static IEnumerable<string> ReadXlsx(string path)
    {
        using var fs = OpenShared(path);
        using var wb = new XLWorkbook(fs);
        foreach (var ws in wb.Worksheets)
            foreach (var cell in ws.CellsUsed())
                yield return cell.GetFormattedString();
    }

    static IEnumerable<string> ReadPptx(string path)
    {
        using var fs = OpenShared(path);
        using var doc = PresentationDocument.Open(fs, false);
        var pres = doc.PresentationPart;
        if (pres == null) yield break;
        foreach (var slidePart in pres.SlideParts)
        {
            var paragraphs = slidePart.Slide?.Descendants<D.Paragraph>() ?? Enumerable.Empty<D.Paragraph>();
            foreach (var p in paragraphs)
                yield return p.InnerText;
        }
    }

    static IEnumerable<string> ReadOpenDocument(string path)
    {
        using var fs = OpenShared(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        var entry = zip.GetEntry("content.xml");
        if (entry == null) yield break;
        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        string xml = sr.ReadToEnd();
        yield return System.Net.WebUtility.HtmlDecode(Regex.Replace(xml, "<[^>]+>", " "));
    }

    static IEnumerable<string> ReadPdf(string path, bool useOcr)
    {
        using var pdf = PdfDocument.Open(path);
        foreach (var page in pdf.GetPages())
        {
            string text = page.Text;
            yield return text;

            if (!useOcr || text.Trim().Length >= 20) continue;
            foreach (var img in page.GetImages())
            {
                if (img.WidthInSamples < 100 || img.HeightInSamples < 100) continue;
                byte[] data = img.TryGetPng(out var png) && png != null ? png : img.RawBytes.ToArray();
                string? ocr = Ocr.Read(data);
                if (!string.IsNullOrWhiteSpace(ocr)) yield return OcrMark + ocr;
            }
        }
    }

    static IEnumerable<string> ReadRaw(string path, string ext)
    {
        using var fs = OpenShared(path);
        if (fs.Length == 0 || fs.Length > MaxFileBytes) yield break;

        var probe = new byte[(int)Math.Min(fs.Length, 8192)];
        int n = fs.Read(probe, 0, probe.Length);
        bool hasNull = Array.IndexOf(probe, (byte)0, 0, n) >= 0;
        bool legacy = LegacyExts.Contains(ext);
        if (hasNull && !legacy) yield break;

        fs.Position = 0;
        var bytes = new byte[fs.Length];
        fs.ReadExactly(bytes);

        if (!hasNull)
        {
            yield return DecodeText(bytes);
            yield break;
        }

        yield return Encoding.GetEncoding(1254).GetString(bytes);
        yield return Encoding.Unicode.GetString(bytes);
        if (bytes.Length > 1)
            yield return Encoding.Unicode.GetString(bytes, 1, bytes.Length - 1);
    }
}

static class Theme
{
    public static bool Dark;
    public static Color Bg, Surface, Surface2, Border, Text, Text2, Muted;
    public static Color Accent, AccentHover, AccentDown, Selected, Hover, Highlight;

    static Color C(int rgb) => Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);

    public static void Set(bool dark)
    {
        Dark = dark;
        if (!dark)
        {
            Bg = C(0xF4F5FA); Surface = C(0xFFFFFF); Surface2 = C(0xF1F3F9); Border = C(0xE2E5EF);
            Text = C(0x1F2430); Text2 = C(0x4B5262); Muted = C(0x8A91A3);
            Accent = C(0x5B5BD6); AccentHover = C(0x4B4BC4); AccentDown = C(0x3F3FB0);
            Selected = C(0xECEEFF); Hover = C(0xF5F6FC); Highlight = C(0xFFE98A);
        }
        else
        {
            Bg = C(0x0F1117); Surface = C(0x181B24); Surface2 = C(0x1F2330); Border = C(0x2A2F3E);
            Text = C(0xE8EAF2); Text2 = C(0xB4B9C9); Muted = C(0x7C8398);
            Accent = C(0x6C6CF5); AccentHover = C(0x7F7FFF); AccentDown = C(0x5B5BE0);
            Selected = C(0x232846); Hover = C(0x1D2130); Highlight = C(0x6B5A00);
        }
    }

    public static Color TypeColor(string type) => type switch
    {
        "Word" => C(0x3B82F6),
        "Excel" => C(0x16A34A),
        "PowerPoint" => C(0xEA580C),
        "PDF" => C(0xDC2626),
        "Image" => C(0x9333EA),
        "Text" => C(0x64748B),
        _ => C(0x0891B2),
    };
}

static class Gfx
{
    public static GraphicsPath Round(RectangleF r, float radius)
    {
        radius = Math.Max(0.5f, Math.Min(radius, Math.Min(r.Width, r.Height) / 2f));
        float d = radius * 2;
        var p = new GraphicsPath();
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

class RoundedPanel : Panel
{
    public Color Fill { get; set; } = Color.White;
    public Color Border { get; set; } = Color.Gray;
    public Color AccentBorder { get; set; } = Color.Blue;
    public bool Accented { get; set; }
    public int Radius { get; set; } = 14;

    public RoundedPanel()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? BackColor);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var path = Gfx.Round(new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), Radius);
        using var fill = new SolidBrush(Fill);
        g.FillPath(fill, path);
        using var pen = new Pen(Accented ? AccentBorder : Border, Accented ? 1.8f : 1f);
        g.DrawPath(pen, path);
    }
}

class ModernButton : Control
{
    public bool Primary { get; set; } = true;
    public bool Danger { get; set; }
    public int Radius { get; set; } = 10;
    bool hover, down;

    public ModernButton()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        SetStyle(ControlStyles.Selectable, false);
        Font = new Font("Segoe UI Semibold", 10f);
        Cursor = Cursors.Hand;
        Size = new Size(100, 40);
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
    protected override void OnTextChanged(EventArgs e) { Invalidate(); base.OnTextChanged(e); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? BackColor);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        Color fill, fg;
        if (Danger)
        {
            fill = hover ? Color.FromArgb(0xD6, 0x3B, 0x40) : Color.FromArgb(0xE5, 0x48, 0x4D);
            fg = Color.White;
        }
        else if (Primary)
        {
            fill = down ? Theme.AccentDown : hover ? Theme.AccentHover : Theme.Accent;
            fg = Color.White;
        }
        else
        {
            fill = hover ? Theme.Hover : Theme.Surface2;
            fg = Theme.Text;
        }

        using var path = Gfx.Round(new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), Radius);
        using var br = new SolidBrush(fill);
        g.FillPath(br, path);
        if (!Primary && !Danger)
        {
            using var pen = new Pen(Theme.Border);
            g.DrawPath(pen, path);
        }
        TextRenderer.DrawText(g, Text, Font, ClientRectangle, fg,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
            TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
    }
}

class Chip : Control
{
    string label = "";
    int count;
    bool selected, hover;
    Color? dot;

    public string Label { get => label; set { label = value; Recalc(); } }
    public int Count { get => count; set { count = value; Recalc(); } }
    public bool Selected { get => selected; set { selected = value; Invalidate(); } }
    public Color? DotColor { get => dot; set { dot = value; Recalc(); } }
    string Caption => $"{label}  {count}";

    public Chip()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        SetStyle(ControlStyles.Selectable, false);
        Font = new Font("Segoe UI Semibold", 9.5f);
        Cursor = Cursors.Hand;
        Height = 34;
        Recalc();
    }

    void Recalc()
    {
        var sz = TextRenderer.MeasureText(Caption, Font);
        Width = sz.Width + (dot.HasValue ? 44 : 32);
        Invalidate();
    }

    protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? BackColor);
        g.SmoothingMode = SmoothingMode.AntiAlias;

        Color fill = selected ? Theme.Accent : hover ? Theme.Hover : Theme.Surface;
        Color border = selected ? Theme.Accent : Theme.Border;
        Color fg = selected ? Color.White : Theme.Text2;

        using var path = Gfx.Round(new RectangleF(0.5f, 0.5f, Width - 1.5f, Height - 1.5f), Height / 2f);
        using (var br = new SolidBrush(fill)) g.FillPath(br, path);
        using (var pen = new Pen(border)) g.DrawPath(pen, path);

        int x = 16;
        if (dot is Color dc)
        {
            using var db = new SolidBrush(selected ? Color.White : dc);
            g.FillEllipse(db, 14, Height / 2 - 4, 8, 8);
            x = 28;
        }
        TextRenderer.DrawText(g, Caption, Font, new Rectangle(x, 0, Width - x - 6, Height), fg,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding |
            TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
    }
}

class ThinProgress : Control
{
    readonly System.Windows.Forms.Timer timer = new() { Interval = 16 };
    float pos = -0.3f;
    bool running;

    public ThinProgress()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        Height = 3;
        timer.Tick += (_, _) =>
        {
            pos += 0.012f;
            if (pos > 1.2f) pos = -0.3f;
            Invalidate();
        };
    }

    public void Start() { running = true; pos = -0.3f; timer.Start(); Invalidate(); }
    public void Stop() { running = false; timer.Stop(); Invalidate(); }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        using (var track = new SolidBrush(Theme.Border)) g.FillRectangle(track, 0, 0, Width, Height);
        if (!running) return;
        int w = (int)(Width * 0.3f);
        int x = (int)(pos * Width);
        using var seg = new SolidBrush(Theme.Accent);
        g.FillRectangle(seg, Math.Max(0, x), 0, Math.Min(w, Width - Math.Max(0, x)) - Math.Max(0, -x), Height);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) timer.Dispose();
        base.Dispose(disposing);
    }
}

class LogoMark : Control
{
    public LogoMark()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                 ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        Size = new Size(44, 44);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.Clear(Parent?.BackColor ?? BackColor);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var path = Gfx.Round(new RectangleF(0, 0, Width - 1, Height - 1), 12))
        using (var br = new SolidBrush(Theme.Accent))
            g.FillPath(br, path);

        using var pen = new Pen(Color.White, 3f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
        float cx = Width * 0.46f, cy = Height * 0.44f, r = Width * 0.2f;
        g.DrawEllipse(pen, cx - r, cy - r, r * 2, r * 2);
        g.DrawLine(pen, cx + r * 0.72f, cy + r * 0.72f, Width * 0.76f, Height * 0.74f);
    }
}

class MainForm : Form
{
    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    readonly List<Hit> allHits = new();
    readonly Dictionary<string, int> typeCounts = new();
    readonly Dictionary<string, Chip> chips = new();
    string? activeFilter;
    string foldedQuery = "";
    string emptyMessage = "";
    int hoverIndex = -1;
    CancellationTokenSource? cts;

    readonly Font fName = new("Segoe UI Semibold", 10.5f);
    readonly Font fSnippet = new("Segoe UI", 9.5f);
    readonly Font fPath = new("Segoe UI", 8.5f);
    readonly Font fBadge = new("Segoe UI Semibold", 8.5f);

    readonly List<Control> bgControls = new();
    readonly List<Control> surfaceControls = new();
    readonly List<(RoundedPanel Panel, TextBox Box)> inputs = new();

    readonly Label titleLabel = new() { Text = "Smart Document Search", Font = new Font("Segoe UI Semibold", 17f), Dock = DockStyle.Top, Height = 34, BackColor = Color.Transparent };
    readonly Label subLabel = new() { Text = "Search inside Word, Excel, PDF, PowerPoint files and images", Font = new Font("Segoe UI", 9.5f), Dock = DockStyle.Top, Height = 22, BackColor = Color.Transparent };
    readonly ModernButton themeBtn = new() { Primary = false, Text = "Dark theme" };
    readonly ModernButton searchBtn = new() { Text = "Search", Font = new Font("Segoe UI Semibold", 11f) };
    readonly ModernButton browseBtn = new() { Primary = false, Text = "Choose folder" };
    readonly CheckBox ocrBox = new()
    {
        Text = "Also read text inside images and scanned PDFs (OCR) - slower",
        Checked = true, Dock = DockStyle.Fill, AutoSize = false, Cursor = Cursors.Hand,
    };
    readonly Chip allChip = new() { Label = "All", Selected = true, Margin = new Padding(0, 0, 8, 0) };
    readonly FlowLayoutPanel chipRow = new() { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = false };
    readonly Panel chipsHost = new() { Dock = DockStyle.Top, Height = 50, Padding = new Padding(24, 2, 24, 12), Visible = false };
    readonly ListView list = new();
    readonly Label emptyLabel = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 11.5f) };
    readonly Label status = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, Padding = new Padding(24, 0, 24, 0), Text = "Ready", BackColor = Color.Transparent };
    readonly ThinProgress progress = new() { Dock = DockStyle.Top };
    RoundedPanel card = null!, resultsCard = null!;
    TextBox queryBox = null!, folderBox = null!;

    public MainForm()
    {
        Theme.Set(false);
        Text = "Smart Document Search";
        Font = new Font("Segoe UI", 10f);
        Width = 1180; Height = 780; MinimumSize = new Size(900, 620);
        StartPosition = FormStartPosition.CenterScreen;
        DoubleBuffered = true;

        var header = new Panel { Dock = DockStyle.Top, Height = 92, Padding = new Padding(24, 20, 24, 6) };
        var headerGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        headerGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 58));
        headerGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        headerGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        headerGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var logo = new LogoMark { Anchor = AnchorStyles.Left };
        var titleBox = new Panel { Dock = DockStyle.Fill };
        titleBox.Controls.Add(subLabel);
        titleBox.Controls.Add(titleLabel);
        themeBtn.Dock = DockStyle.Fill;
        themeBtn.Margin = new Padding(0, 6, 0, 6);
        headerGrid.Controls.Add(logo, 0, 0);
        headerGrid.Controls.Add(titleBox, 1, 0);
        headerGrid.Controls.Add(themeBtn, 2, 0);
        header.Controls.Add(headerGrid);

        var cardHost = new Panel { Dock = DockStyle.Top, Height = 196, Padding = new Padding(24, 6, 24, 12) };
        card = new RoundedPanel { Dock = DockStyle.Fill, Padding = new Padding(18, 16, 18, 8), Radius = 16 };
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3 };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));

        var searchInput = MakeInput(out queryBox, "Type a word to search for", 12.5f);
        searchInput.Margin = new Padding(0, 0, 12, 10);
        searchBtn.Dock = DockStyle.Fill;
        searchBtn.Margin = new Padding(0, 0, 0, 10);

        var folderInput = MakeInput(out folderBox, "Folder to scan", 10f);
        folderInput.Margin = new Padding(0, 0, 12, 8);
        folderBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        browseBtn.Dock = DockStyle.Fill;
        browseBtn.Margin = new Padding(0, 0, 0, 8);

        grid.Controls.Add(searchInput, 0, 0);
        grid.Controls.Add(searchBtn, 1, 0);
        grid.Controls.Add(folderInput, 0, 1);
        grid.Controls.Add(browseBtn, 1, 1);
        grid.Controls.Add(ocrBox, 0, 2);
        grid.SetColumnSpan(ocrBox, 2);
        card.Controls.Add(grid);
        cardHost.Controls.Add(card);

        chipRow.Controls.Add(allChip);
        allChip.Click += (_, _) => SetFilter(null);
        chipsHost.Controls.Add(chipRow);

        var resultsHost = new Panel { Dock = DockStyle.Fill, Padding = new Padding(24, 0, 24, 8) };
        resultsCard = new RoundedPanel { Dock = DockStyle.Fill, Padding = new Padding(8), Radius = 16 };
        SetupList();
        resultsCard.Controls.Add(list);
        resultsCard.Controls.Add(emptyLabel);
        resultsHost.Controls.Add(resultsCard);

        var statusPanel = new Panel { Dock = DockStyle.Bottom, Height = 44 };
        statusPanel.Controls.Add(status);
        statusPanel.Controls.Add(progress);

        Controls.Add(resultsHost);
        Controls.Add(statusPanel);
        Controls.Add(chipsHost);
        Controls.Add(cardHost);
        Controls.Add(header);

        bgControls.AddRange(new Control[] { header, headerGrid, titleBox, cardHost, chipsHost, chipRow, resultsHost, statusPanel });
        surfaceControls.AddRange(new Control[] { grid, ocrBox });

        browseBtn.Click += (_, _) =>
        {
            using var dlg = new FolderBrowserDialog { SelectedPath = folderBox.Text };
            if (dlg.ShowDialog() == DialogResult.OK) folderBox.Text = dlg.SelectedPath;
        };
        searchBtn.Click += async (_, _) => await OnSearchClick();
        queryBox.KeyDown += async (_, e) =>
        {
            if (e.KeyCode != Keys.Enter) return;
            e.SuppressKeyPress = true;
            if (cts == null) await OnSearchClick();
        };
        themeBtn.Click += (_, _) => { Theme.Set(!Theme.Dark); ApplyTheme(); };

        emptyMessage = "Type a word and press Search to begin\nSearches inside Word, Excel, PDF, PowerPoint, images and more";
        UpdateEmpty();
        ApplyTheme();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        SetTitleBarDark(Theme.Dark);
    }

    void SetTitleBarDark(bool dark)
    {
        if (!IsHandleCreated) return;
        int v = dark ? 1 : 0;
        try { DwmSetWindowAttribute(Handle, 20, ref v, sizeof(int)); } catch { }
    }

    RoundedPanel MakeInput(out TextBox tb, string placeholder, float fontSize)
    {
        var panel = new RoundedPanel { Dock = DockStyle.Fill, Radius = 10 };
        var box = new TextBox
        {
            BorderStyle = BorderStyle.None,
            PlaceholderText = placeholder,
            Font = new Font("Segoe UI", fontSize),
        };
        panel.Controls.Add(box);
        void Place()
        {
            box.Left = 14;
            box.Width = Math.Max(10, panel.Width - 28);
            box.Top = Math.Max(0, (panel.Height - box.Height) / 2);
        }
        panel.Resize += (_, _) => Place();
        box.Enter += (_, _) => { panel.Accented = true; panel.Invalidate(); };
        box.Leave += (_, _) => { panel.Accented = false; panel.Invalidate(); };
        inputs.Add((panel, box));
        tb = box;
        return panel;
    }

    void SetupList()
    {
        list.Dock = DockStyle.Fill;
        list.View = View.Details;
        list.HeaderStyle = ColumnHeaderStyle.None;
        list.FullRowSelect = true;
        list.MultiSelect = false;
        list.HideSelection = false;
        list.BorderStyle = BorderStyle.None;
        list.OwnerDraw = true;
        list.Columns.Add("", 400);
        list.SmallImageList = new ImageList { ImageSize = new Size(1, 70) };
        typeof(ListView).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance)!
            .SetValue(list, true);

        list.DrawItem += (_, _) => { };
        list.DrawSubItem += DrawRow;
        list.ClientSizeChanged += (_, _) =>
        {
            if (list.Columns.Count > 0) list.Columns[0].Width = Math.Max(100, list.ClientSize.Width);
        };
        list.MouseMove += (_, e) =>
        {
            int idx = list.GetItemAt(e.X, e.Y)?.Index ?? -1;
            if (idx == hoverIndex) return;
            int old = hoverIndex;
            hoverIndex = idx;
            if (old >= 0 && old < list.Items.Count) list.Invalidate(list.Items[old].Bounds);
            if (idx >= 0) list.Invalidate(list.Items[idx].Bounds);
            list.Cursor = idx >= 0 ? Cursors.Hand : Cursors.Default;
        };
        list.MouseLeave += (_, _) =>
        {
            hoverIndex = -1;
            list.Cursor = Cursors.Default;
            list.Invalidate();
        };
        list.DoubleClick += (_, _) => OpenSelected();
        list.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) OpenSelected(); };

        var menu = new ContextMenuStrip();
        menu.Items.Add("Open file", null, (_, _) => OpenSelected());
        menu.Items.Add("Show in folder", null, (_, _) => ShowInFolder());
        list.ContextMenuStrip = menu;
    }

    void ApplyTheme()
    {
        BackColor = Theme.Bg;
        ForeColor = Theme.Text;
        foreach (var c in bgControls) c.BackColor = Theme.Bg;
        foreach (var c in surfaceControls) c.BackColor = Theme.Surface;

        card.Fill = Theme.Surface; card.Border = Theme.Border;
        resultsCard.Fill = Theme.Surface; resultsCard.Border = Theme.Border;
        foreach (var (panel, box) in inputs)
        {
            panel.Fill = Theme.Surface2; panel.Border = Theme.Border; panel.AccentBorder = Theme.Accent;
            box.BackColor = Theme.Surface2; box.ForeColor = Theme.Text;
        }

        list.BackColor = Theme.Surface;
        emptyLabel.BackColor = Theme.Surface; emptyLabel.ForeColor = Theme.Muted;
        titleLabel.ForeColor = Theme.Text;
        subLabel.ForeColor = Theme.Muted;
        status.ForeColor = Theme.Text2;
        ocrBox.ForeColor = Theme.Text2;
        themeBtn.Text = Theme.Dark ? "Light theme" : "Dark theme";

        SetTitleBarDark(Theme.Dark);
        Invalidate(true);
        list.Invalidate();
    }

    static int TextW(Graphics g, string s, Font f)
    {
        var flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine;
        var big = new Size(int.MaxValue, int.MaxValue);
        return TextRenderer.MeasureText(g, s + "|", f, big, flags).Width
             - TextRenderer.MeasureText(g, "|", f, big, flags).Width;
    }

    void DrawHighlighted(Graphics g, string text, Font font, Rectangle rect, Color color)
    {
        var flags = TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine |
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter;
        int idx = foldedQuery.Length > 0
            ? SearchEngine.Fold(text).IndexOf(foldedQuery, StringComparison.Ordinal) : -1;
        if (idx < 0)
        {
            TextRenderer.DrawText(g, text, font, rect, color, flags | TextFormatFlags.EndEllipsis);
            return;
        }

        string a = text[..idx];
        string m = text.Substring(idx, foldedQuery.Length);
        string b = text[(idx + foldedQuery.Length)..];
        int wa = TextW(g, a, font), wm = TextW(g, m, font);

        var state = g.Save();
        g.SetClip(rect);

        var hl = new Rectangle(rect.X + wa - 2, rect.Y + 1, wm + 4, Math.Max(1, rect.Height - 2));
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (var path = Gfx.Round(hl, 4))
        using (var br = new SolidBrush(Theme.Highlight))
            g.FillPath(br, path);

        void Part(string s, int x)
        {
            if (s.Length == 0) return;
            TextRenderer.DrawText(g, s, font, new Rectangle(x, rect.Y, Math.Max(1, rect.Right - x), rect.Height), color, flags);
        }
        Part(a, rect.X);
        Part(m, rect.X + wa);
        Part(b, rect.X + wa + wm);
        g.Restore(state);
    }

    void DrawRow(object? sender, DrawListViewSubItemEventArgs e)
    {
        if (e.Item?.Tag is not Hit h) return;
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;

        var b = e.Item.GetBounds(ItemBoundsPortion.Entire);
        var r = new Rectangle(0, b.Y, list.ClientSize.Width, b.Height);

        Color bg = e.Item.Selected ? Theme.Selected : e.ItemIndex == hoverIndex ? Theme.Hover : Theme.Surface;
        using (var br = new SolidBrush(bg)) g.FillRectangle(br, r);
        if (e.Item.Selected)
            using (var ab = new SolidBrush(Theme.Accent)) g.FillRectangle(ab, r.X, r.Y + 12, 3, r.Height - 24);

        Color tc = Theme.TypeColor(h.Type);
        var badge = new Rectangle(r.X + 18, r.Y + (r.Height - 28) / 2, 92, 28);
        using (var path = Gfx.Round(badge, 8))
        using (var br = new SolidBrush(Color.FromArgb(Theme.Dark ? 60 : 36, tc)))
            g.FillPath(br, path);
        TextRenderer.DrawText(g, h.Type, fBadge, badge, tc,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
            TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);

        int x = r.X + 128;
        int w = Math.Max(50, r.Right - x - 20);
        DrawHighlighted(g, h.FileName, fName, new Rectangle(x, r.Y + 8, w, 20), Theme.Text);
        DrawHighlighted(g, h.Snippet, fSnippet, new Rectangle(x, r.Y + 29, w, 18), Theme.Text2);
        TextRenderer.DrawText(g, h.Folder, fPath, new Rectangle(x, r.Y + 48, w, 16), Theme.Muted,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.PathEllipsis |
            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);

        using var pen = new Pen(Theme.Border);
        g.DrawLine(pen, r.X + 18, r.Bottom - 1, r.Right - 18, r.Bottom - 1);
    }

    void UpdateEmpty()
    {
        bool empty = list.Items.Count == 0;
        emptyLabel.Text = emptyMessage;
        emptyLabel.Visible = empty;
        list.Visible = !empty;
    }

    ListViewItem MakeItem(Hit h) => new(h.FileName) { Tag = h };

    void RegisterHit(Hit h)
    {
        allHits.Add(h);
        typeCounts[h.Type] = typeCounts.GetValueOrDefault(h.Type) + 1;

        if (!chips.TryGetValue(h.Type, out var chip))
        {
            string t = h.Type;
            chip = new Chip { Label = t, DotColor = Theme.TypeColor(t), Margin = new Padding(0, 0, 8, 0) };
            chip.Click += (_, _) => SetFilter(t);
            chips[t] = chip;
            chipRow.Controls.Add(chip);
        }
        chip.Count = typeCounts[h.Type];
        allChip.Count = allHits.Count;

        if (activeFilter == null || activeFilter == h.Type)
            list.Items.Add(MakeItem(h));
        UpdateEmpty();
    }

    void ResetResults()
    {
        allHits.Clear();
        typeCounts.Clear();
        foreach (var c in chips.Values) { chipRow.Controls.Remove(c); c.Dispose(); }
        chips.Clear();
        activeFilter = null;
        allChip.Selected = true;
        allChip.Count = 0;
        hoverIndex = -1;
        list.Items.Clear();
    }

    void SetFilter(string? type)
    {
        activeFilter = type;
        allChip.Selected = type == null;
        foreach (var kv in chips) kv.Value.Selected = kv.Key == type;

        list.BeginUpdate();
        list.Items.Clear();
        hoverIndex = -1;
        foreach (var h in allHits)
            if (type == null || h.Type == type) list.Items.Add(MakeItem(h));
        list.EndUpdate();
        UpdateEmpty();
    }

    async Task OnSearchClick()
    {
        if (cts != null)
        {
            cts.Cancel();
            status.Text = "Stopping...";
            return;
        }

        string query = queryBox.Text.Trim();
        string folder = folderBox.Text.Trim();
        bool useOcr = ocrBox.Checked;
        if (query.Length == 0) { MessageBox.Show("Please enter a word to search for."); return; }
        if (!Directory.Exists(folder)) { MessageBox.Show("Please choose a valid folder."); return; }

        ResetResults();
        chipsHost.Visible = true;
        foldedQuery = SearchEngine.Fold(query);
        emptyMessage = "Searching...";
        UpdateEmpty();

        searchBtn.Text = "Stop";
        searchBtn.Danger = true;
        progress.Start();
        cts = new CancellationTokenSource();
        var token = cts.Token;

        int scanned = 0, found = 0, unreadable = 0;
        bool ocrOk = false;

        var onHit = new Progress<Hit>(h =>
        {
            found++;
            RegisterHit(h);
        });
        var onScan = new Progress<string>(p =>
        {
            scanned++;
            status.Text = $"Scanning...  {scanned} files  •  {found} results  •  {Path.GetFileName(p)}";
        });

        try
        {
            await Task.Run(() =>
            {
                if (useOcr) ocrOk = Ocr.Available;

                var options = new EnumerationOptions
                {
                    RecurseSubdirectories = true,
                    IgnoreInaccessible = true,
                };
                var files = new System.IO.Enumeration.FileSystemEnumerable<string>(
                    folder,
                    (ref System.IO.Enumeration.FileSystemEntry entry) => entry.ToFullPath(),
                    options)
                {
                    ShouldIncludePredicate = (ref System.IO.Enumeration.FileSystemEntry entry) => !entry.IsDirectory,
                    ShouldRecursePredicate = (ref System.IO.Enumeration.FileSystemEntry entry) =>
                        !SearchEngine.SkipDirs.Contains(entry.FileName.ToString()),
                };

                foreach (string path in files)
                {
                    token.ThrowIfCancellationRequested();
                    string name = Path.GetFileName(path);
                    if (name.StartsWith("~$")) continue;

                    ((IProgress<string>)onScan).Report(path);

                    using var fileCts = CancellationTokenSource.CreateLinkedTokenSource(token);
                    fileCts.CancelAfter(TimeSpan.FromSeconds(15));
                    var fileTask = Task.Run(() =>
                        SearchEngine.SearchFile(path, folder, foldedQuery, useOcr, fileCts.Token));

                    bool finished = fileTask.Wait(20000, token);
                    if (!finished)
                    {
                        fileCts.Cancel();
                        unreadable++;
                        continue;
                    }

                    var (snippet, failed) = fileTask.Result;
                    if (failed) unreadable++;
                    if (snippet != null)
                    {
                        string ext = Path.GetExtension(path).ToLowerInvariant();
                        ((IProgress<Hit>)onHit).Report(
                            new Hit(SearchEngine.TypeOf(ext), name, Path.GetDirectoryName(path)!, snippet));
                    }
                }
            }, token);
            status.Text = BuildFinalStatus("Done", scanned, found, unreadable, useOcr, ocrOk);
        }
        catch (OperationCanceledException)
        {
            status.Text = BuildFinalStatus("Stopped", scanned, found, unreadable, useOcr, ocrOk);
        }
        finally
        {
            cts?.Dispose();
            cts = null;
            progress.Stop();
            searchBtn.Danger = false;
            searchBtn.Text = "Search";
            emptyMessage = allHits.Count == 0
                ? "No results found\nTry a different word or choose another folder"
                : "";
            UpdateEmpty();
        }
    }

    static string BuildFinalStatus(string prefix, int scanned, int found, int unreadable, bool useOcr, bool ocrOk)
    {
        string s = $"{prefix}  •  {scanned} files scanned  •  {found} results";
        if (unreadable > 0) s += $"  •  {unreadable} files skipped (encrypted, corrupted, locked or took too long)";
        if (useOcr && !ocrOk) s += $"  •  OCR did not run: {Ocr.Error}";
        return s;
    }

    string? SelectedPath() =>
        list.SelectedItems.Count > 0 && list.SelectedItems[0].Tag is Hit h
            ? Path.Combine(h.Folder, h.FileName) : null;

    void OpenSelected()
    {
        string? path = SelectedPath();
        if (path == null) return;
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show("Could not open file: " + ex.Message); }
    }

    void ShowInFolder()
    {
        string? path = SelectedPath();
        if (path != null)
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{path}\""));
    }
}
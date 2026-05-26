using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using XmlDiffTool.Models;

namespace XmlDiffTool.Services
{
    public class HtmlReportService
    {
        private const string ApplicationFolderName = "XmlDiffTool";
        private const string AssetFolderName = "ReportAssets";

        public string AssetDirectory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            ApplicationFolderName,
            AssetFolderName);

        public void EnsureAssets()
        {
            Directory.CreateDirectory(AssetDirectory);
            WriteAsset("bootstrap.min.css", BootstrapCss);
            WriteAsset("bootstrap.bundle.min.js", BootstrapJs);
            WriteAsset("xml-diff-report.css", ReportCss);
            WriteAsset("xml-diff-report.js", ReportJs);
        }

        public void SaveReport(string reportPath, IReadOnlyCollection<XmlDifferenceNode> roots, string? leftFilePath, string? rightFilePath, bool ignoreCase)
        {
            EnsureAssets();
            File.WriteAllText(reportPath, BuildHtml(roots, leftFilePath, rightFilePath, ignoreCase), Encoding.UTF8);
        }

        private string BuildHtml(IReadOnlyCollection<XmlDifferenceNode> roots, string? leftFilePath, string? rightFilePath, bool ignoreCase)
        {
            var total = CountDifferenceRows(roots);
            var leftOnly = CountNodes(roots, node => node.IsRightMissing);
            var rightOnly = CountNodes(roots, node => node.IsLeftMissing);
            var generatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            var builder = new StringBuilder();
            builder.AppendLine("<!doctype html>");
            builder.AppendLine("<html lang=\"en\">");
            builder.AppendLine("<head>");
            builder.AppendLine("  <meta charset=\"utf-8\">");
            builder.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
            builder.AppendLine("  <title>XML Diff Report</title>");
            builder.AppendLine($"  <link rel=\"stylesheet\" href=\"{ToFileUri(Path.Combine(AssetDirectory, "bootstrap.min.css"))}\">");
            builder.AppendLine($"  <link rel=\"stylesheet\" href=\"{ToFileUri(Path.Combine(AssetDirectory, "xml-diff-report.css"))}\">");
            builder.AppendLine("</head>");
            builder.AppendLine("<body>");
            builder.AppendLine("  <main class=\"container-fluid py-4\">");
            builder.AppendLine("    <header class=\"report-header mb-3\">");
            builder.AppendLine("      <div>");
            builder.AppendLine("        <h1>XML Diff Report</h1>");
            builder.AppendLine($"        <div class=\"text-muted small\">Generated {Encode(generatedAt)} &middot; Ignore case: {Encode(ignoreCase ? "Yes" : "No")} &middot; List order ignored</div>");
            builder.AppendLine("      </div>");
            builder.AppendLine("    </header>");
            builder.AppendLine("    <section class=\"summary-grid mb-3\">");
            builder.AppendLine($"      <div><span>Total differences</span><strong>{total}</strong></div>");
            builder.AppendLine($"      <div><span>Left only</span><strong>{leftOnly}</strong></div>");
            builder.AppendLine($"      <div><span>Right only</span><strong>{rightOnly}</strong></div>");
            builder.AppendLine("    </section>");
            builder.AppendLine("    <section class=\"file-grid mb-3\">");
            builder.AppendLine($"      <div><span>Left XML</span><code>{Encode(leftFilePath ?? string.Empty)}</code></div>");
            builder.AppendLine($"      <div><span>Right XML</span><code>{Encode(rightFilePath ?? string.Empty)}</code></div>");
            builder.AppendLine("    </section>");
            builder.AppendLine("    <section class=\"toolbar sticky-top py-2 mb-3\">");
            builder.AppendLine("      <div class=\"form-check form-check-inline\">");
            builder.AppendLine("        <input class=\"form-check-input diff-filter\" type=\"checkbox\" id=\"showLeftOnly\" data-filter=\"left-only\" checked>");
            builder.AppendLine("        <label class=\"form-check-label\" for=\"showLeftOnly\">Show left only element</label>");
            builder.AppendLine("      </div>");
            builder.AppendLine("      <div class=\"form-check form-check-inline\">");
            builder.AppendLine("        <input class=\"form-check-input diff-filter\" type=\"checkbox\" id=\"showRightOnly\" data-filter=\"right-only\" checked>");
            builder.AppendLine("        <label class=\"form-check-label\" for=\"showRightOnly\">Show right only element</label>");
            builder.AppendLine("      </div>");
            builder.AppendLine("      <div class=\"form-check form-check-inline\">");
            builder.AppendLine("        <input class=\"form-check-input\" type=\"checkbox\" id=\"showLineNumbers\" data-action=\"line-numbers\" checked>");
            builder.AppendLine("        <label class=\"form-check-label\" for=\"showLineNumbers\">Show line numbers</label>");
            builder.AppendLine("      </div>");
            builder.AppendLine("      <button class=\"btn btn-sm btn-outline-secondary\" type=\"button\" data-action=\"expand\">Expand all</button>");
            builder.AppendLine("      <button class=\"btn btn-sm btn-outline-secondary\" type=\"button\" data-action=\"collapse\">Collapse all</button>");
            builder.AppendLine("    </section>");
            builder.AppendLine("    <section class=\"diff-list\">");

            if (roots.Count == 0)
            {
                builder.AppendLine("      <div class=\"alert alert-success\">No differences found.</div>");
            }
            else
            {
                foreach (var root in roots)
                {
                    AppendNode(builder, root, 0, leftFilePath, rightFilePath);
                }
            }

            builder.AppendLine("    </section>");
            builder.AppendLine("  </main>");
            builder.AppendLine($"  <script src=\"{ToFileUri(Path.Combine(AssetDirectory, "bootstrap.bundle.min.js"))}\"></script>");
            builder.AppendLine($"  <script src=\"{ToFileUri(Path.Combine(AssetDirectory, "xml-diff-report.js"))}\"></script>");
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
            return builder.ToString();
        }

        private static void AppendNode(StringBuilder builder, XmlDifferenceNode node, int depth, string? leftFilePath, string? rightFilePath)
        {
            var id = $"node-{Guid.NewGuid():N}";
            var sideClass = node.IsLeftMissing ? "right-only" : node.IsRightMissing ? "left-only" : "changed";
            var kindClass = node.Kind.ToString().ToLowerInvariant();
            var hasChildren = node.Children.Count > 0;
            var directRows = node.Children.Where(child => !child.HasChildren).ToList();
            var childSections = node.Children.Where(child => child.HasChildren).ToList();
            var differenceCount = CountDifferenceRows(new[] { node });
            var indent = Math.Min(depth, 8);

            builder.AppendLine($"      <article class=\"diff-section {sideClass} {kindClass}\" data-side=\"{sideClass}\" style=\"--depth:{indent}\">");
            builder.AppendLine("        <div class=\"section-heading\">");

            if (hasChildren)
            {
                builder.AppendLine($"          <button class=\"toggle\" type=\"button\" data-bs-toggle=\"collapse\" data-bs-target=\"#{id}\" aria-expanded=\"true\" aria-controls=\"{id}\">-</button>");
            }
            else
            {
                builder.AppendLine("          <span class=\"toggle-spacer\"></span>");
            }

            builder.AppendLine($"          <h2>{Encode(node.Path)}</h2>");
            builder.AppendLine($"          <span class=\"count-badge\">{differenceCount}</span>");
            builder.AppendLine("        </div>");

            if (hasChildren)
            {
                builder.AppendLine($"        <div id=\"{id}\" class=\"collapse show section-body\">");
            }

            if (directRows.Count > 0 || !hasChildren)
            {
                AppendDifferenceTable(builder, node, directRows.Count > 0 ? directRows : new List<XmlDifferenceNode> { node }, leftFilePath, rightFilePath);
            }

            foreach (var childSection in childSections)
            {
                AppendNode(builder, childSection, depth + 1, leftFilePath, rightFilePath);
            }

            if (hasChildren)
            {
                builder.AppendLine("        </div>");
            }

            builder.AppendLine("      </article>");
        }

        private static void AppendDifferenceTable(StringBuilder builder, XmlDifferenceNode owner, IReadOnlyCollection<XmlDifferenceNode> rows, string? leftFilePath, string? rightFilePath)
        {
            builder.AppendLine("          <table class=\"diff-table\">");
            builder.AppendLine("            <thead>");
            builder.AppendLine("              <tr>");
            builder.AppendLine("                <th>Property</th>");
            builder.AppendLine($"                <th class=\"left-head\">{Encode(GetFileLabel(leftFilePath, "Left"))}</th>");
            builder.AppendLine($"                <th class=\"right-head\">{Encode(GetFileLabel(rightFilePath, "Right"))}</th>");
            builder.AppendLine("              </tr>");
            builder.AppendLine("            </thead>");
            builder.AppendLine("            <tbody>");

            foreach (var row in rows)
            {
                var sideClass = row.IsLeftMissing ? "right-only" : row.IsRightMissing ? "left-only" : "changed";
                builder.AppendLine($"              <tr class=\"{sideClass}\" data-side=\"{sideClass}\">");
                builder.AppendLine($"                <td class=\"property-name\">{BuildPropertyLabel(owner, row)}</td>");
                builder.AppendLine($"                <td class=\"left-value\">{BuildValueCell(row.LeftValue, row.IsLeftMissing, row.LeftLineNumber)}</td>");
                builder.AppendLine($"                <td class=\"right-value\">{BuildValueCell(row.RightValue, row.IsRightMissing, row.RightLineNumber)}</td>");
                builder.AppendLine("              </tr>");
            }

            builder.AppendLine("            </tbody>");
            builder.AppendLine("          </table>");
        }

        private static string BuildPropertyLabel(XmlDifferenceNode owner, XmlDifferenceNode row)
        {
            var propertyName = Encode(GetPropertyName(owner, row));
            var note = GetOnlySideNote(row);
            return string.IsNullOrEmpty(note)
                ? propertyName
                : $"{propertyName} <span class=\"side-note\">{Encode(note)}</span>";
        }

        private static string BuildValueCell(string? value, bool isMissing, int? lineNumber)
        {
            return $"<div class=\"line-number\">{Encode(DisplayLineNumber(lineNumber, isMissing))}</div><pre>{Encode(DisplayValue(value, isMissing))}</pre>";
        }

        private static string GetPropertyName(XmlDifferenceNode owner, XmlDifferenceNode row)
        {
            if (row.Kind == XmlDifferenceKind.Attribute)
            {
                return row.Name.TrimStart('@');
            }

            if (row.Kind == XmlDifferenceKind.Value)
            {
                return "Value";
            }

            if (row.Path.StartsWith(owner.Path + "/", StringComparison.Ordinal))
            {
                return row.Path[(owner.Path.Length + 1)..];
            }

            return row.Name;
        }

        private static string GetOnlySideNote(XmlDifferenceNode row)
        {
            if (row.IsRightMissing)
            {
                return "(left only)";
            }

            if (row.IsLeftMissing)
            {
                return "(right only)";
            }

            return string.Empty;
        }

        private static string GetFileLabel(string? filePath, string fallback)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return fallback;
            }

            return Path.GetFileName(filePath);
        }

        private static string DisplayValue(string? value, bool isMissing)
        {
            return isMissing ? "(missing)" : value ?? string.Empty;
        }

        private static string DisplayLineNumber(int? lineNumber, bool isMissing)
        {
            if (isMissing)
            {
                return "Line -";
            }

            return lineNumber is null ? "Line unknown" : $"Line {lineNumber.Value}";
        }

        private static int CountNodes(IEnumerable<XmlDifferenceNode> nodes)
        {
            return nodes.Sum(node => 1 + CountNodes(node.Children));
        }

        private static int CountDifferenceRows(IEnumerable<XmlDifferenceNode> nodes)
        {
            return nodes.Sum(node => node.Children.Count == 0 ? 1 : CountDifferenceRows(node.Children));
        }

        private static int CountNodes(IEnumerable<XmlDifferenceNode> nodes, Func<XmlDifferenceNode, bool> predicate)
        {
            return nodes.Sum(node => (predicate(node) ? 1 : 0) + CountNodes(node.Children, predicate));
        }

        private static string Encode(string value)
        {
            return WebUtility.HtmlEncode(value);
        }

        private static string ToFileUri(string path)
        {
            return new Uri(path).AbsoluteUri;
        }

        private void WriteAsset(string fileName, string content)
        {
            var path = Path.Combine(AssetDirectory, fileName);
            File.WriteAllText(path, content, Encoding.UTF8);
        }

        private const string BootstrapCss = ":root{--bs-body-font-family:Segoe UI,Arial,sans-serif;--bs-body-color:#212529;--bs-body-bg:#f8f9fa;--bs-border-color:#dee2e6;--bs-primary:#0d6efd;--bs-danger:#dc3545;--bs-success:#198754;--bs-secondary:#6c757d}"
            + "*,::after,::before{box-sizing:border-box}body{margin:0;font-family:var(--bs-body-font-family);font-size:1rem;color:var(--bs-body-color);background:var(--bs-body-bg)}h1{margin:0;font-size:1.65rem}code{font-family:Cascadia Mono,Consolas,monospace;color:#222;word-break:break-word}.container-fluid{width:100%;padding-right:1rem;padding-left:1rem;margin-right:auto;margin-left:auto}.py-4{padding-top:1.5rem;padding-bottom:1.5rem}.py-2{padding-top:.5rem;padding-bottom:.5rem}.mb-3{margin-bottom:1rem}.small{font-size:.875rem}.text-muted{color:#6c757d}.sticky-top{position:sticky;top:0;z-index:1020}.form-check{display:inline-block;min-height:1.5rem;padding-left:1.5em;margin-right:1rem}.form-check-input{width:1em;height:1em;margin-top:.25em;margin-left:-1.5em;vertical-align:top}.form-check-label{cursor:pointer}.btn{display:inline-block;font-weight:400;line-height:1.5;text-align:center;text-decoration:none;vertical-align:middle;cursor:pointer;user-select:none;background-color:transparent;border:1px solid transparent;padding:.375rem .75rem;font-size:1rem;border-radius:.375rem}.btn-sm{padding:.25rem .5rem;font-size:.875rem;border-radius:.25rem}.btn-outline-secondary{color:#6c757d;border-color:#6c757d}.btn-outline-secondary:hover{color:#fff;background-color:#6c757d}.badge{display:inline-block;padding:.35em .65em;font-size:.75em;font-weight:700;line-height:1;text-align:center;white-space:nowrap;border-radius:.375rem}.text-bg-light{color:#000;background-color:#f8f9fa}.alert{position:relative;padding:1rem;border:1px solid transparent;border-radius:.375rem}.alert-success{color:#0f5132;background-color:#d1e7dd;border-color:#badbcc}.collapse:not(.show){display:none}";

        private const string BootstrapJs = "(()=>{document.addEventListener(\"click\",event=>{const trigger=event.target.closest(\"[data-bs-toggle='collapse']\");if(!trigger)return;const target=document.querySelector(trigger.getAttribute(\"data-bs-target\"));if(!target)return;target.classList.toggle(\"show\");trigger.setAttribute(\"aria-expanded\",target.classList.contains(\"show\"));trigger.textContent=target.classList.contains(\"show\")?\"-\":\"+\";});})();";

        private const string ReportCss = ".report-header{display:flex;align-items:flex-end;justify-content:space-between}.summary-grid,.file-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:.75rem}.summary-grid>div,.file-grid>div{background:#fff;border:1px solid var(--bs-border-color);border-radius:8px;padding:.85rem}.summary-grid span,.file-grid span{display:block;color:#6c757d;font-size:.78rem;text-transform:uppercase}.summary-grid strong{font-size:1.4rem}.toolbar{background:rgba(248,249,250,.97);border-bottom:1px solid var(--bs-border-color);box-shadow:0 2px 8px rgba(0,0,0,.06)}.diff-list{display:flex;flex-direction:column;gap:.9rem}.diff-section{margin-left:calc(var(--depth)*1rem);background:#fff;border:1px solid var(--bs-border-color);border-radius:8px;padding:1rem;box-shadow:0 1px 3px rgba(0,0,0,.05)}.section-heading{display:flex;align-items:center;gap:.65rem;border-bottom:2px solid #dbe7f3;padding-bottom:.75rem;margin-bottom:1rem}.section-heading h2{margin:0;color:#0d2f5f;font-size:1.2rem;font-weight:700}.count-badge{display:inline-flex;align-items:center;justify-content:center;min-width:1.65rem;height:1.65rem;padding:0 .5rem;border-radius:999px;background:#bfe4ff;color:#0d4e85;font-weight:700}.toggle,.toggle-spacer{width:1.75rem;height:1.75rem;flex:0 0 1.75rem}.toggle{border:1px solid var(--bs-border-color);border-radius:6px;background:#fff;cursor:pointer}.toggle-spacer{display:inline-block}.section-body{display:flex;flex-direction:column;gap:.85rem}.diff-table{width:100%;border-collapse:collapse;table-layout:fixed}.diff-table th{position:sticky;top:3rem;z-index:5;background:#eaf1f7;color:#23384f;text-align:left;font-weight:600;padding:.75rem;border-bottom:1px solid #cbd8e6}.diff-table th:first-child{width:18%}.diff-table td{padding:.75rem;border-bottom:1px solid #d9e1ea;vertical-align:top}.property-name{background:#fff;color:#0b223f;font-weight:600}.side-note{color:#d00000;font-size:.86rem;font-weight:700;white-space:nowrap}.left-value{background:#ffd6d6;color:#b00020}.right-value{background:#c9f7d8;color:#005c2f}.left-head{background:#eef3f8}.right-head{background:#eef3f8}.line-number{display:inline-block;margin-bottom:.35rem;padding:.1rem .4rem;border-radius:4px;background:rgba(255,255,255,.72);color:#334155;font-family:Cascadia Mono,Consolas,monospace;font-size:.78rem;font-weight:700}.hide-line-numbers .line-number{display:none}.diff-table pre{margin:0;white-space:pre-wrap;overflow-wrap:anywhere;font-family:Cascadia Mono,Consolas,monospace}.is-hidden-by-filter{display:none}@media(max-width:720px){.diff-section{margin-left:0;padding:.75rem}.diff-table{table-layout:auto}.diff-table th{top:4.25rem}.diff-table th:first-child{width:auto}.toolbar .btn{margin-top:.5rem}}";

        private const string ReportJs = "(()=>{const filters=document.querySelectorAll(\".diff-filter\");const nodes=[...document.querySelectorAll(\".diff-section,tr[data-side]\")];const lineToggle=document.querySelector(\"[data-action='line-numbers']\");function applyFilters(){const showLeft=document.querySelector(\"[data-filter='left-only']\").checked;const showRight=document.querySelector(\"[data-filter='right-only']\").checked;for(const node of nodes){const side=node.dataset.side;node.classList.toggle(\"is-hidden-by-filter\",(side===\"left-only\"&&!showLeft)||(side===\"right-only\"&&!showRight));}}function applyLineNumbers(){document.body.classList.toggle(\"hide-line-numbers\",lineToggle&&!lineToggle.checked);}for(const filter of filters)filter.addEventListener(\"change\",applyFilters);lineToggle?.addEventListener(\"change\",applyLineNumbers);document.addEventListener(\"click\",event=>{const action=event.target.closest(\"button[data-action]\")?.dataset.action;if(!action)return;const show=action===\"expand\";for(const collapse of document.querySelectorAll(\".collapse\"))collapse.classList.toggle(\"show\",show);for(const toggle of document.querySelectorAll(\".toggle\")){toggle.setAttribute(\"aria-expanded\",show);toggle.textContent=show?\"-\":\"+\";}});applyFilters();applyLineNumbers();})();";
    }
}

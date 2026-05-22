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
            var total = CountNodes(roots);
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
                    AppendNode(builder, root, 0);
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

        private static void AppendNode(StringBuilder builder, XmlDifferenceNode node, int depth)
        {
            var id = $"node-{Guid.NewGuid():N}";
            var sideClass = node.IsLeftMissing ? "right-only" : node.IsRightMissing ? "left-only" : "changed";
            var kindClass = node.Kind.ToString().ToLowerInvariant();
            var hasChildren = node.Children.Count > 0;
            var hasVisibleValues = ShouldShowValues(node);
            var indent = Math.Min(depth, 8);

            builder.AppendLine($"      <article class=\"diff-node {sideClass} {kindClass}\" data-side=\"{sideClass}\" style=\"--depth:{indent}\">");
            builder.AppendLine("        <div class=\"diff-row\">");

            if (hasChildren)
            {
                builder.AppendLine($"          <button class=\"toggle\" type=\"button\" data-bs-toggle=\"collapse\" data-bs-target=\"#{id}\" aria-expanded=\"true\" aria-controls=\"{id}\">-</button>");
            }
            else
            {
                builder.AppendLine("          <span class=\"toggle-spacer\"></span>");
            }

            builder.AppendLine("          <div class=\"diff-main\">");
            builder.AppendLine($"            <div class=\"diff-title{(hasVisibleValues ? string.Empty : " mb-0")}\"><span class=\"badge text-bg-light\">{Encode(node.Kind.ToString())}</span><code>{Encode(node.Path)}</code></div>");

            if (hasVisibleValues)
            {
                builder.AppendLine("            <div class=\"diff-values\">");
                builder.AppendLine($"              <div class=\"value-pane left\"><span>Left</span><pre>{Encode(DisplayValue(node.LeftValue, node.IsLeftMissing))}</pre></div>");
                builder.AppendLine($"              <div class=\"value-pane right\"><span>Right</span><pre>{Encode(DisplayValue(node.RightValue, node.IsRightMissing))}</pre></div>");
                builder.AppendLine("            </div>");
            }

            builder.AppendLine("          </div>");
            builder.AppendLine("        </div>");

            if (hasChildren)
            {
                builder.AppendLine($"        <div id=\"{id}\" class=\"collapse show diff-children\">");
                foreach (var child in node.Children)
                {
                    AppendNode(builder, child, depth + 1);
                }

                builder.AppendLine("        </div>");
            }

            builder.AppendLine("      </article>");
        }

        private static string DisplayValue(string? value, bool isMissing)
        {
            return isMissing ? "(missing)" : value ?? string.Empty;
        }

        private static bool ShouldShowValues(XmlDifferenceNode node)
        {
            return node.IsLeftMissing
                || node.IsRightMissing
                || !string.IsNullOrEmpty(node.LeftValue)
                || !string.IsNullOrEmpty(node.RightValue);
        }

        private static int CountNodes(IEnumerable<XmlDifferenceNode> nodes)
        {
            return nodes.Sum(node => 1 + CountNodes(node.Children));
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

        private const string ReportCss = ".report-header{display:flex;align-items:flex-end;justify-content:space-between}.summary-grid,.file-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:.75rem}.summary-grid>div,.file-grid>div{background:#fff;border:1px solid var(--bs-border-color);border-radius:8px;padding:.85rem}.summary-grid span,.file-grid span{display:block;color:#6c757d;font-size:.78rem;text-transform:uppercase}.summary-grid strong{font-size:1.4rem}.toolbar{background:rgba(248,249,250,.95);border-bottom:1px solid var(--bs-border-color)}.diff-list{display:flex;flex-direction:column;gap:.65rem}.diff-node{margin-left:calc(var(--depth)*1rem)}.diff-row{display:flex;gap:.5rem;background:#fff;border:1px solid var(--bs-border-color);border-left:4px solid #0d6efd;border-radius:8px;padding:.75rem}.diff-node.left-only>.diff-row{border-left-color:#dc3545}.diff-node.right-only>.diff-row{border-left-color:#198754}.toggle,.toggle-spacer{width:1.75rem;height:1.75rem;flex:0 0 1.75rem}.toggle{border:1px solid var(--bs-border-color);border-radius:6px;background:#fff;cursor:pointer}.toggle-spacer{display:inline-block}.diff-main{min-width:0;flex:1}.diff-title{display:flex;align-items:center;gap:.5rem;margin-bottom:.6rem}.diff-title.mb-0{margin-bottom:0}.diff-values{display:grid;grid-template-columns:1fr 1fr;gap:.75rem}.value-pane{min-width:0;border:1px solid var(--bs-border-color);border-radius:6px;overflow:hidden;background:#fbfbfc}.value-pane span{display:block;padding:.35rem .5rem;font-size:.75rem;font-weight:700;color:#6c757d;border-bottom:1px solid var(--bs-border-color)}.value-pane pre{margin:0;padding:.6rem;min-height:2.4rem;white-space:pre-wrap;overflow-wrap:anywhere;font-family:Cascadia Mono,Consolas,monospace}.diff-children{margin-top:.6rem;display:flex;flex-direction:column;gap:.6rem}.is-hidden-by-filter{display:none}@media(max-width:720px){.diff-values{grid-template-columns:1fr}.diff-node{margin-left:0}.toolbar .btn{margin-top:.5rem}}";

        private const string ReportJs = "(()=>{const filters=document.querySelectorAll(\".diff-filter\");const nodes=[...document.querySelectorAll(\".diff-node\")];function applyFilters(){const showLeft=document.querySelector(\"[data-filter='left-only']\").checked;const showRight=document.querySelector(\"[data-filter='right-only']\").checked;for(const node of nodes){const side=node.dataset.side;node.classList.toggle(\"is-hidden-by-filter\",(side===\"left-only\"&&!showLeft)||(side===\"right-only\"&&!showRight));}}for(const filter of filters)filter.addEventListener(\"change\",applyFilters);document.addEventListener(\"click\",event=>{const action=event.target.closest(\"[data-action]\")?.dataset.action;if(!action)return;const show=action===\"expand\";for(const collapse of document.querySelectorAll(\".collapse\"))collapse.classList.toggle(\"show\",show);for(const toggle of document.querySelectorAll(\".toggle\")){toggle.setAttribute(\"aria-expanded\",show);toggle.textContent=show?\"-\":\"+\";}});applyFilters();})();";
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using XmlDiffTool.Models;

namespace XmlDiffTool.Services
{
    public class XmlComparer
    {
        public IReadOnlyCollection<ParameterDifference> Compare(string leftPath, string rightPath, IProgress<int>? progress = null)
        {
            progress?.Report(0);

            var leftParameters = LoadParameters(leftPath, progress, 0, 45);
            var rightParameters = LoadParameters(rightPath, progress, 45, 90);

            var keys = new HashSet<string>(leftParameters.Keys);
            keys.UnionWith(rightParameters.Keys);

            var orderedKeys = new List<string>(keys);
            orderedKeys.Sort(System.StringComparer.Ordinal);

            var differences = new List<ParameterDifference>(orderedKeys.Count);
            for (var index = 0; index < orderedKeys.Count; index++)
            {
                var key = orderedKeys[index];
                var (title, displayName) = ParseKey(key);
                leftParameters.TryGetValue(key, out var leftValue);
                rightParameters.TryGetValue(key, out var rightValue);

                differences.Add(new ParameterDifference(key, title, displayName, leftValue, rightValue));
                ReportScaledProgress(progress, 90, 100, index + 1, orderedKeys.Count);
            }

            progress?.Report(100);
            return differences;
        }

        private static (string Title, string DisplayName) ParseKey(string key)
        {
            var attributeMarkerIndex = key.LastIndexOf("[@", System.StringComparison.Ordinal);
            if (attributeMarkerIndex >= 0)
            {
                var attributeTagPath = key[..attributeMarkerIndex];
                var attributeStart = attributeMarkerIndex + 2;
                var attributeEnd = key.IndexOf(']', attributeStart);
                var attributeName = attributeEnd > attributeStart
                    ? key[attributeStart..attributeEnd]
                    : key[attributeStart..];

                return (GetTagName(attributeTagPath), attributeName);
            }

            var valueMarkerIndex = key.LastIndexOf("[#", System.StringComparison.Ordinal);
            var valueTagPath = valueMarkerIndex >= 0 ? key[..valueMarkerIndex] : key;
            var currentTagName = GetTagName(valueTagPath);
            var parentTagName = GetParentTagName(valueTagPath);

            return (parentTagName, currentTagName);
        }

        private static string GetParentTagName(string path)
        {
            var lastSlashIndex = path.LastIndexOf('/');
            if (lastSlashIndex < 0)
            {
                return GetTagName(path);
            }

            return GetTagName(path[..lastSlashIndex]);
        }

        private static string GetTagName(string path)
        {
            var lastSlashIndex = path.LastIndexOf('/');
            var tagSegment = lastSlashIndex >= 0 ? path[(lastSlashIndex + 1)..] : path;
            var tagNameEnd = tagSegment.IndexOf('[');
            var tagName = tagNameEnd >= 0 ? tagSegment[..tagNameEnd] : tagSegment;

            return string.IsNullOrWhiteSpace(tagName) ? "Unknown" : tagName;
        }

        private static Dictionary<string, string> LoadParameters(string path, IProgress<int>? progress, int progressStart, int progressEnd)
        {
            var settings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreProcessingInstructions = true
            };

            var values = new Dictionary<string, string>();
            var elementStack = new Stack<ElementContext>();

            using var stream = File.OpenRead(path);
            using var reader = XmlReader.Create(stream, settings);
            progress?.Report(progressStart);
            var lastReportedProgress = progressStart;

            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        var elementContext = CreateElementContext(reader, elementStack, values);
                        if (reader.IsEmptyElement)
                        {
                            values[$"{elementContext.Path}[#text]"] = string.Empty;
                            continue;
                        }

                        elementStack.Push(elementContext);
                        break;

                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                    case XmlNodeType.Whitespace:
                        if (elementStack.Count > 0)
                        {
                            elementStack.Peek().AppendText(reader.Value);
                        }

                        break;

                    case XmlNodeType.EndElement:
                        if (elementStack.Count == 0)
                        {
                            continue;
                        }

                        var completedElement = elementStack.Pop();
                        if (!completedElement.HasChildElements)
                        {
                            values[$"{completedElement.Path}[#text]"] = completedElement.GetValue();
                        }

                        break;
                }

                ReportStreamProgress(progress, progressStart, progressEnd, stream, lastReportedProgress, forceReport: false, out lastReportedProgress);
            }

            ReportStreamProgress(progress, progressStart, progressEnd, stream, lastReportedProgress, forceReport: true, out _);
            return values;
        }


        private static void ReportScaledProgress(IProgress<int>? progress, int start, int end, int completed, int total)
        {
            if (progress is null || total <= 0)
            {
                return;
            }

            var percent = start + (int)(((long)(end - start) * completed) / total);
            progress.Report(percent);
        }

        private static void ReportStreamProgress(IProgress<int>? progress, int start, int end, FileStream stream, int lastReportedProgress, bool forceReport, out int updatedLastReportedProgress)
        {
            updatedLastReportedProgress = lastReportedProgress;
            if (progress is null)
            {
                return;
            }

            if (stream.Length <= 0)
            {
                if (forceReport && lastReportedProgress != end)
                {
                    progress.Report(end);
                    updatedLastReportedProgress = end;
                }

                return;
            }

            var percent = start + (int)(((long)(end - start) * stream.Position) / stream.Length);
            if (forceReport)
            {
                percent = end;
            }

            if (percent != lastReportedProgress)
            {
                progress.Report(percent);
                updatedLastReportedProgress = percent;
            }
        }

        private static ElementContext CreateElementContext(XmlReader reader, Stack<ElementContext> elementStack, IDictionary<string, string> values)
        {
            string path;
            if (elementStack.Count == 0)
            {
                path = reader.LocalName;
            }
            else
            {
                var parentContext = elementStack.Peek();
                parentContext.HasChildElements = true;
                var childIndex = parentContext.GetNextChildIndex(reader.LocalName);
                path = $"{parentContext.Path}/{reader.LocalName}[{childIndex}]";
            }

            if (reader.HasAttributes)
            {
                while (reader.MoveToNextAttribute())
                {
                    values[$"{path}[@{reader.LocalName}]"] = reader.Value;
                }

                reader.MoveToElement();
            }

            return new ElementContext(path);
        }

        private sealed class ElementContext
        {
            private readonly Dictionary<string, int> _childNameCounts = new();
            private StringBuilder? _textBuilder;

            public ElementContext(string path)
            {
                Path = path;
            }

            public string Path { get; }

            public bool HasChildElements { get; set; }

            public int GetNextChildIndex(string childName)
            {
                _childNameCounts.TryGetValue(childName, out var currentIndex);
                _childNameCounts[childName] = currentIndex + 1;
                return currentIndex;
            }

            public void AppendText(string value)
            {
                _textBuilder ??= new StringBuilder();
                _textBuilder.Append(value);
            }

            public string GetValue()
            {
                return _textBuilder?.ToString() ?? string.Empty;
            }
        }
    }
}

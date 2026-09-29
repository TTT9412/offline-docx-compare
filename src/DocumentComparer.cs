using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OfflineDocxCompare
{
    public static class DocumentComparer
    {
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace Xml = "http://www.w3.org/XML/1998/namespace";
        private static readonly Regex HeaderFooterPattern = new Regex(
            @"^word/(header|footer)\d+\.xml$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly HashSet<string> AllowedHighlightColors = new HashSet<string>(
            new[]
            {
                "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue",
                "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
                "darkGray", "lightGray", "black", "white"
            },
            StringComparer.OrdinalIgnoreCase);

        public static CompareResult Compare(
            string oldPath,
            string newPath,
            string outputPath,
            string highlightColor,
            IProgress<int> progress)
        {
            ValidateInputs(oldPath, newPath, outputPath, highlightColor);

            string canonicalColor = AllowedHighlightColors
                .First(value => value.Equals(highlightColor, StringComparison.OrdinalIgnoreCase));
            string outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            Directory.CreateDirectory(outputDirectory);

            string temporaryPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileName(outputPath) + ".tmp-" + Guid.NewGuid().ToString("N"));

            CompareResult result = new CompareResult();
            result.OutputPath = Path.GetFullPath(outputPath);

            try
            {
                using (FileStream oldFile = OpenReadOnly(oldPath))
                using (FileStream newFile = OpenReadOnly(newPath))
                using (ZipArchive oldArchive = new ZipArchive(oldFile, ZipArchiveMode.Read, false))
                using (ZipArchive newArchive = new ZipArchive(newFile, ZipArchiveMode.Read, false))
                {
                    EnsureDocxPackage(oldArchive, "旧版本");
                    EnsureDocxPackage(newArchive, "新版本");

                    Dictionary<string, ZipArchiveEntry> oldEntries = oldArchive.Entries
                        .ToDictionary(entry => NormalizePartName(entry.FullName), StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, byte[]> modifiedParts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                    List<ZipArchiveEntry> newStoryEntries = newArchive.Entries
                        .Where(entry => IsStoryPart(NormalizePartName(entry.FullName)))
                        .ToList();

                    int totalStages = Math.Max(1, newStoryEntries.Count + 1);
                    int completedStages = 0;

                    foreach (ZipArchiveEntry newEntry in newStoryEntries)
                    {
                        string partName = NormalizePartName(newEntry.FullName);
                        byte[] newXmlBytes = ReadEntryBytes(newEntry, 64 * 1024 * 1024);
                        XDocument newDocument = LoadXmlSafely(newXmlBytes, partName);
                        XDocument oldDocument = null;
                        ZipArchiveEntry oldEntry;
                        if (oldEntries.TryGetValue(partName, out oldEntry))
                        {
                            oldDocument = LoadXmlSafely(ReadEntryBytes(oldEntry, 64 * 1024 * 1024), partName);
                        }

                        bool changed = CompareStory(oldDocument, newDocument, canonicalColor, result);
                        if (changed)
                        {
                            modifiedParts[partName] = SaveXml(newDocument);
                        }

                        result.ComparedParts++;
                        completedStages++;
                        Report(progress, completedStages * 90 / totalStages);
                    }

                    HashSet<string> newPartNames = new HashSet<string>(
                        newArchive.Entries.Select(entry => NormalizePartName(entry.FullName)),
                        StringComparer.OrdinalIgnoreCase);
                    List<Tuple<string, string>> orphanedDeletions = new List<Tuple<string, string>>();
                    foreach (KeyValuePair<string, ZipArchiveEntry> oldPart in oldEntries)
                    {
                        if (IsStoryPart(oldPart.Key) && !newPartNames.Contains(oldPart.Key))
                        {
                            XDocument removedDocument = LoadXmlSafely(
                                ReadEntryBytes(oldPart.Value, 64 * 1024 * 1024),
                                oldPart.Key);
                            foreach (ParagraphInfo removedParagraph in ExtractParagraphs(removedDocument))
                            {
                                if (!string.IsNullOrWhiteSpace(removedParagraph.Text))
                                {
                                    orphanedDeletions.Add(Tuple.Create(oldPart.Key, removedParagraph.Text));
                                    result.DeletedParagraphs++;
                                    result.DeletedCharactersShown += removedParagraph.Text.Length;
                                }
                            }
                        }
                    }

                    if (orphanedDeletions.Count > 0)
                    {
                        const string mainPartName = "word/document.xml";
                        byte[] mainDocumentBytes;
                        if (!modifiedParts.TryGetValue(mainPartName, out mainDocumentBytes))
                        {
                            ZipArchiveEntry mainEntry = newArchive.Entries.First(entry =>
                                NormalizePartName(entry.FullName).Equals(mainPartName, StringComparison.OrdinalIgnoreCase));
                            mainDocumentBytes = ReadEntryBytes(mainEntry, 64 * 1024 * 1024);
                        }
                        XDocument mainDocument = LoadXmlSafely(mainDocumentBytes, mainPartName);
                        AppendOrphanedDeletionSection(mainDocument, orphanedDeletions);
                        modifiedParts[mainPartName] = SaveXml(mainDocument);
                    }

                    using (FileStream outputFile = new FileStream(
                        temporaryPath,
                        FileMode.CreateNew,
                        FileAccess.ReadWrite,
                        FileShare.None))
                    using (ZipArchive outputArchive = new ZipArchive(outputFile, ZipArchiveMode.Create, false))
                    {
                        int copied = 0;
                        int entryCount = Math.Max(1, newArchive.Entries.Count);
                        foreach (ZipArchiveEntry sourceEntry in newArchive.Entries)
                        {
                            CopyEntry(sourceEntry, outputArchive, modifiedParts);
                            copied++;
                            Report(progress, 90 + copied * 9 / entryCount);
                        }
                    }
                }

                if (File.Exists(outputPath))
                {
                    try
                    {
                        File.Replace(temporaryPath, outputPath, null);
                    }
                    catch (IOException)
                    {
                        File.Delete(outputPath);
                        File.Move(temporaryPath, outputPath);
                    }
                    catch (PlatformNotSupportedException)
                    {
                        File.Delete(outputPath);
                        File.Move(temporaryPath, outputPath);
                    }
                }
                else
                {
                    File.Move(temporaryPath, outputPath);
                }
                Report(progress, 100);
                return result;
            }
            catch
            {
                TryDelete(temporaryPath);
                throw;
            }
        }

        private static void ValidateInputs(
            string oldPath,
            string newPath,
            string outputPath,
            string highlightColor)
        {
            if (string.IsNullOrWhiteSpace(oldPath) || !File.Exists(oldPath))
            {
                throw new ArgumentException("请选择有效的旧版本 DOCX 文件。");
            }
            if (string.IsNullOrWhiteSpace(newPath) || !File.Exists(newPath))
            {
                throw new ArgumentException("请选择有效的新版本 DOCX 文件。");
            }
            if (!string.Equals(Path.GetExtension(oldPath), ".docx", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetExtension(newPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("当前版本仅支持 .docx 文件，不支持旧式 .doc 文件。");
            }
            if (string.IsNullOrWhiteSpace(outputPath) ||
                !string.Equals(Path.GetExtension(outputPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("输出文件必须使用 .docx 扩展名。");
            }

            string oldFullPath = Path.GetFullPath(oldPath);
            string newFullPath = Path.GetFullPath(newPath);
            string outputFullPath = Path.GetFullPath(outputPath);
            if (outputFullPath.Equals(oldFullPath, StringComparison.OrdinalIgnoreCase) ||
                outputFullPath.Equals(newFullPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("输出文件不能覆盖旧版本或新版本源文件。");
            }
            if (!AllowedHighlightColors.Contains(highlightColor))
            {
                throw new ArgumentException("所选高亮颜色不是 Word 支持的标准高亮色。");
            }
        }

        private static FileStream OpenReadOnly(string path)
        {
            return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        }

        private static void EnsureDocxPackage(ZipArchive archive, string label)
        {
            bool hasContentTypes = archive.Entries.Any(entry =>
                NormalizePartName(entry.FullName).Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
            bool hasDocument = archive.Entries.Any(entry =>
                NormalizePartName(entry.FullName).Equals("word/document.xml", StringComparison.OrdinalIgnoreCase));
            if (!hasContentTypes || !hasDocument)
            {
                throw new InvalidDataException(label + "不是有效的 DOCX 文档，或文档已经损坏。");
            }
        }

        private static bool IsStoryPart(string partName)
        {
            return partName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase) ||
                   partName.Equals("word/footnotes.xml", StringComparison.OrdinalIgnoreCase) ||
                   partName.Equals("word/endnotes.xml", StringComparison.OrdinalIgnoreCase) ||
                   HeaderFooterPattern.IsMatch(partName);
        }

        private static string NormalizePartName(string name)
        {
            return name.Replace('\\', '/').TrimStart('/');
        }

        private static bool CompareStory(
            XDocument oldDocument,
            XDocument newDocument,
            string highlightColor,
            CompareResult result)
        {
            List<ParagraphInfo> newParagraphs = ExtractParagraphs(newDocument);
            List<ParagraphInfo> oldParagraphs = oldDocument == null
                ? new List<ParagraphInfo>()
                : ExtractParagraphs(oldDocument);

            List<ParagraphPair> alignment = AlignParagraphs(oldParagraphs, newParagraphs);
            bool changed = false;

            for (int pairIndex = 0; pairIndex < alignment.Count; pairIndex++)
            {
                ParagraphPair pair = alignment[pairIndex];
                if (pair.OldIndex >= 0 && pair.NewIndex >= 0)
                {
                    ParagraphInfo oldParagraph = oldParagraphs[pair.OldIndex];
                    ParagraphInfo newParagraph = newParagraphs[pair.NewIndex];
                    if (oldParagraph.Text.Equals(newParagraph.Text, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    ParagraphDiff paragraphDiff = ComputeParagraphDiff(oldParagraph.Text, newParagraph.Text);
                    int highlighted = CountTrue(paragraphDiff.ChangedMask);
                    ApplyHighlight(newParagraph.Element, paragraphDiff.ChangedMask, highlightColor);
                    InsertDeletionAnnotations(newParagraph.Element, paragraphDiff.Deletions);
                    result.DeletedCharactersShown += paragraphDiff.Deletions.Sum(annotation => annotation.DeletedText.Length);
                    result.ChangedParagraphs++;
                    result.ReplacedParagraphs++;
                    result.HighlightedCharacters += highlighted;
                    changed = true;
                }
                else if (pair.NewIndex >= 0)
                {
                    ParagraphInfo newParagraph = newParagraphs[pair.NewIndex];
                    bool[] changedMask = Enumerable.Repeat(true, newParagraph.Text.Length).ToArray();
                    ApplyHighlight(newParagraph.Element, changedMask, highlightColor);
                    result.ChangedParagraphs++;
                    result.InsertedParagraphs++;
                    result.HighlightedCharacters += newParagraph.Text.Length;
                    changed = true;
                }
                else if (pair.OldIndex >= 0)
                {
                    ParagraphInfo oldParagraph = oldParagraphs[pair.OldIndex];
                    if (!string.IsNullOrWhiteSpace(oldParagraph.Text))
                    {
                        XElement previousParagraph = FindNeighborParagraph(alignment, newParagraphs, pairIndex, -1);
                        XElement nextParagraph = FindNeighborParagraph(alignment, newParagraphs, pairIndex, 1);
                        InsertDeletedParagraph(newDocument, previousParagraph, nextParagraph, oldParagraph.Text);
                        result.DeletedCharactersShown += oldParagraph.Text.Length;
                    }
                    result.DeletedParagraphs++;
                    changed = true;
                }
            }

            return changed;
        }

        private static List<ParagraphInfo> ExtractParagraphs(XDocument document)
        {
            List<ParagraphInfo> paragraphs = new List<ParagraphInfo>();
            if (document.Root == null)
            {
                return paragraphs;
            }

            foreach (XElement paragraph in document.Descendants(W + "p"))
            {
                List<XElement> textNodes = paragraph
                    .Descendants(W + "t")
                    .Where(text => text.Ancestors(W + "p").FirstOrDefault() == paragraph)
                    .ToList();
                string textValue = string.Concat(textNodes.Select(text => text.Value));
                paragraphs.Add(new ParagraphInfo
                {
                    Element = paragraph,
                    Text = textValue,
                    Fingerprint = BuildFingerprint(textValue)
                });
            }
            return paragraphs;
        }

        private static List<ParagraphPair> AlignParagraphs(
            List<ParagraphInfo> oldParagraphs,
            List<ParagraphInfo> newParagraphs)
        {
            List<Tuple<int, int>> anchors = FindUniqueAnchors(oldParagraphs, newParagraphs);
            List<ParagraphPair> result = new List<ParagraphPair>();
            int oldStart = 0;
            int newStart = 0;

            foreach (Tuple<int, int> anchor in anchors)
            {
                result.AddRange(AlignGap(oldParagraphs, oldStart, anchor.Item1, newParagraphs, newStart, anchor.Item2));
                result.Add(new ParagraphPair(anchor.Item1, anchor.Item2));
                oldStart = anchor.Item1 + 1;
                newStart = anchor.Item2 + 1;
            }

            result.AddRange(AlignGap(
                oldParagraphs,
                oldStart,
                oldParagraphs.Count,
                newParagraphs,
                newStart,
                newParagraphs.Count));
            return result;
        }

        private static List<Tuple<int, int>> FindUniqueAnchors(
            List<ParagraphInfo> oldParagraphs,
            List<ParagraphInfo> newParagraphs)
        {
            Dictionary<string, List<int>> oldPositions = BuildPositionMap(oldParagraphs);
            Dictionary<string, List<int>> newPositions = BuildPositionMap(newParagraphs);
            List<Tuple<int, int>> candidates = new List<Tuple<int, int>>();

            foreach (KeyValuePair<string, List<int>> item in oldPositions)
            {
                List<int> newIndexes;
                if (item.Value.Count == 1 &&
                    newPositions.TryGetValue(item.Key, out newIndexes) &&
                    newIndexes.Count == 1 &&
                    item.Key.Length > 0)
                {
                    candidates.Add(Tuple.Create(item.Value[0], newIndexes[0]));
                }
            }

            candidates.Sort((left, right) => left.Item1.CompareTo(right.Item1));
            return LongestIncreasingSubsequence(candidates);
        }

        private static Dictionary<string, List<int>> BuildPositionMap(List<ParagraphInfo> paragraphs)
        {
            Dictionary<string, List<int>> positions = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (int index = 0; index < paragraphs.Count; index++)
            {
                List<int> indexes;
                if (!positions.TryGetValue(paragraphs[index].Text, out indexes))
                {
                    indexes = new List<int>();
                    positions[paragraphs[index].Text] = indexes;
                }
                indexes.Add(index);
            }
            return positions;
        }

        private static List<Tuple<int, int>> LongestIncreasingSubsequence(List<Tuple<int, int>> candidates)
        {
            if (candidates.Count == 0)
            {
                return new List<Tuple<int, int>>();
            }

            int[] tails = new int[candidates.Count];
            int[] previous = Enumerable.Repeat(-1, candidates.Count).ToArray();
            int length = 0;

            for (int index = 0; index < candidates.Count; index++)
            {
                int value = candidates[index].Item2;
                int low = 0;
                int high = length;
                while (low < high)
                {
                    int middle = (low + high) / 2;
                    if (candidates[tails[middle]].Item2 < value)
                    {
                        low = middle + 1;
                    }
                    else
                    {
                        high = middle;
                    }
                }

                if (low > 0)
                {
                    previous[index] = tails[low - 1];
                }
                tails[low] = index;
                if (low == length)
                {
                    length++;
                }
            }

            List<Tuple<int, int>> result = new List<Tuple<int, int>>();
            int cursor = tails[length - 1];
            while (cursor >= 0)
            {
                result.Add(candidates[cursor]);
                cursor = previous[cursor];
            }
            result.Reverse();
            return result;
        }

        private static List<ParagraphPair> AlignGap(
            List<ParagraphInfo> oldParagraphs,
            int oldStart,
            int oldEnd,
            List<ParagraphInfo> newParagraphs,
            int newStart,
            int newEnd)
        {
            int oldCount = oldEnd - oldStart;
            int newCount = newEnd - newStart;
            List<ParagraphPair> result = new List<ParagraphPair>();

            if (oldCount == 0)
            {
                for (int index = newStart; index < newEnd; index++)
                {
                    result.Add(new ParagraphPair(-1, index));
                }
                return result;
            }
            if (newCount == 0)
            {
                for (int index = oldStart; index < oldEnd; index++)
                {
                    result.Add(new ParagraphPair(index, -1));
                }
                return result;
            }

            long cellCount = (long)(oldCount + 1) * (long)(newCount + 1);
            if (cellCount > 350000)
            {
                int paired = Math.Min(oldCount, newCount);
                for (int index = 0; index < paired; index++)
                {
                    result.Add(new ParagraphPair(oldStart + index, newStart + index));
                }
                for (int index = paired; index < oldCount; index++)
                {
                    result.Add(new ParagraphPair(oldStart + index, -1));
                }
                for (int index = paired; index < newCount; index++)
                {
                    result.Add(new ParagraphPair(-1, newStart + index));
                }
                return result;
            }

            double[] previousCosts = new double[newCount + 1];
            double[] currentCosts = new double[newCount + 1];
            byte[] directions = new byte[(oldCount + 1) * (newCount + 1)];
            for (int column = 0; column <= newCount; column++)
            {
                previousCosts[column] = column;
            }

            for (int row = 1; row <= oldCount; row++)
            {
                currentCosts[0] = row;
                for (int column = 1; column <= newCount; column++)
                {
                    ParagraphInfo oldParagraph = oldParagraphs[oldStart + row - 1];
                    ParagraphInfo newParagraph = newParagraphs[newStart + column - 1];
                    double substitutionCost;
                    if (oldParagraph.Text.Equals(newParagraph.Text, StringComparison.Ordinal))
                    {
                        substitutionCost = 0.0;
                    }
                    else
                    {
                        double similarity = FingerprintSimilarity(oldParagraph, newParagraph);
                        substitutionCost = similarity >= 0.12 ? 1.15 - (0.90 * similarity) : 2.05;
                    }

                    double diagonal = previousCosts[column - 1] + substitutionCost;
                    double deletion = previousCosts[column] + 1.0;
                    double insertion = currentCosts[column - 1] + 1.0;
                    int offset = row * (newCount + 1) + column;

                    if (diagonal <= deletion && diagonal <= insertion)
                    {
                        currentCosts[column] = diagonal;
                        directions[offset] = 0;
                    }
                    else if (deletion <= insertion)
                    {
                        currentCosts[column] = deletion;
                        directions[offset] = 1;
                    }
                    else
                    {
                        currentCosts[column] = insertion;
                        directions[offset] = 2;
                    }
                }

                double[] swap = previousCosts;
                previousCosts = currentCosts;
                currentCosts = swap;
            }

            int traceRow = oldCount;
            int traceColumn = newCount;
            while (traceRow > 0 || traceColumn > 0)
            {
                if (traceRow == 0)
                {
                    result.Add(new ParagraphPair(-1, newStart + traceColumn - 1));
                    traceColumn--;
                    continue;
                }
                if (traceColumn == 0)
                {
                    result.Add(new ParagraphPair(oldStart + traceRow - 1, -1));
                    traceRow--;
                    continue;
                }

                byte direction = directions[traceRow * (newCount + 1) + traceColumn];
                if (direction == 0)
                {
                    result.Add(new ParagraphPair(oldStart + traceRow - 1, newStart + traceColumn - 1));
                    traceRow--;
                    traceColumn--;
                }
                else if (direction == 1)
                {
                    result.Add(new ParagraphPair(oldStart + traceRow - 1, -1));
                    traceRow--;
                }
                else
                {
                    result.Add(new ParagraphPair(-1, newStart + traceColumn - 1));
                    traceColumn--;
                }
            }

            result.Reverse();
            return result;
        }

        private static HashSet<int> BuildFingerprint(string text)
        {
            string normalized = Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim();
            HashSet<int> fingerprint = new HashSet<int>();
            if (normalized.Length == 0)
            {
                return fingerprint;
            }
            if (normalized.Length == 1)
            {
                fingerprint.Add(normalized[0]);
                return fingerprint;
            }

            for (int index = 0; index < normalized.Length - 1; index++)
            {
                fingerprint.Add((normalized[index] << 16) ^ normalized[index + 1]);
            }
            return fingerprint;
        }

        private static double FingerprintSimilarity(ParagraphInfo left, ParagraphInfo right)
        {
            if (left.Text.Equals(right.Text, StringComparison.Ordinal))
            {
                return 1.0;
            }
            int maximumLength = Math.Max(left.Text.Length, right.Text.Length);
            int minimumLength = Math.Min(left.Text.Length, right.Text.Length);
            if (maximumLength == 0 || (double)minimumLength / maximumLength < 0.20)
            {
                return 0.0;
            }
            if (left.Fingerprint.Count == 0 || right.Fingerprint.Count == 0)
            {
                return 0.0;
            }

            HashSet<int> smaller = left.Fingerprint.Count <= right.Fingerprint.Count
                ? left.Fingerprint
                : right.Fingerprint;
            HashSet<int> larger = ReferenceEquals(smaller, left.Fingerprint)
                ? right.Fingerprint
                : left.Fingerprint;
            int intersection = 0;
            foreach (int value in smaller)
            {
                if (larger.Contains(value))
                {
                    intersection++;
                }
            }
            return (2.0 * intersection) / (left.Fingerprint.Count + right.Fingerprint.Count);
        }

        private static ParagraphDiff ComputeParagraphDiff(string oldText, string newText)
        {
            bool[] changed = new bool[newText.Length];
            List<DeletionAnnotation> deletions = new List<DeletionAnnotation>();
            List<TextToken> oldTokens = Tokenize(oldText);
            List<TextToken> newTokens = Tokenize(newText);

            int prefix = 0;
            while (prefix < oldTokens.Count &&
                   prefix < newTokens.Count &&
                   oldTokens[prefix].Value.Equals(newTokens[prefix].Value, StringComparison.Ordinal))
            {
                prefix++;
            }

            int suffix = 0;
            while (suffix < oldTokens.Count - prefix &&
                   suffix < newTokens.Count - prefix &&
                   oldTokens[oldTokens.Count - 1 - suffix].Value.Equals(
                       newTokens[newTokens.Count - 1 - suffix].Value,
                       StringComparison.Ordinal))
            {
                suffix++;
            }

            int oldMiddleCount = oldTokens.Count - prefix - suffix;
            int newMiddleCount = newTokens.Count - prefix - suffix;
            if (newMiddleCount <= 0)
            {
                AddDeletionAnnotation(
                    deletions,
                    NewOffsetForToken(newTokens, prefix, newText.Length),
                    TextForTokenRange(oldText, oldTokens, prefix, oldMiddleCount));
                return new ParagraphDiff { ChangedMask = changed, Deletions = deletions };
            }
            if (oldMiddleCount <= 0)
            {
                MarkTokens(changed, newTokens, prefix, newMiddleCount);
                return new ParagraphDiff { ChangedMask = changed, Deletions = deletions };
            }

            long cells = (long)(oldMiddleCount + 1) * (long)(newMiddleCount + 1);
            if (cells > 2000000)
            {
                MarkTokens(changed, newTokens, prefix, newMiddleCount);
                AddDeletionAnnotation(
                    deletions,
                    NewOffsetForToken(newTokens, prefix, newText.Length),
                    TextForTokenRange(oldText, oldTokens, prefix, oldMiddleCount));
                return new ParagraphDiff { ChangedMask = changed, Deletions = deletions };
            }

            int[] previous = new int[newMiddleCount + 1];
            int[] current = new int[newMiddleCount + 1];
            byte[] directions = new byte[(oldMiddleCount + 1) * (newMiddleCount + 1)];

            for (int row = 1; row <= oldMiddleCount; row++)
            {
                for (int column = 1; column <= newMiddleCount; column++)
                {
                    string oldValue = oldTokens[prefix + row - 1].Value;
                    string newValue = newTokens[prefix + column - 1].Value;
                    int offset = row * (newMiddleCount + 1) + column;
                    if (oldValue.Equals(newValue, StringComparison.Ordinal))
                    {
                        current[column] = previous[column - 1] + 1;
                        directions[offset] = 0;
                    }
                    else if (previous[column] >= current[column - 1])
                    {
                        current[column] = previous[column];
                        directions[offset] = 1;
                    }
                    else
                    {
                        current[column] = current[column - 1];
                        directions[offset] = 2;
                    }
                }

                int[] swap = previous;
                previous = current;
                current = swap;
                Array.Clear(current, 0, current.Length);
            }

            List<Tuple<int, int>> matches = new List<Tuple<int, int>>();
            int traceRow = oldMiddleCount;
            int traceColumn = newMiddleCount;
            while (traceRow > 0 && traceColumn > 0)
            {
                byte direction = directions[traceRow * (newMiddleCount + 1) + traceColumn];
                if (direction == 0)
                {
                    matches.Add(Tuple.Create(traceRow - 1, traceColumn - 1));
                    traceRow--;
                    traceColumn--;
                }
                else if (direction == 1)
                {
                    traceRow--;
                }
                else
                {
                    traceColumn--;
                }
            }
            matches.Reverse();

            int oldCursor = 0;
            int newCursor = 0;
            foreach (Tuple<int, int> match in matches)
            {
                if (match.Item1 > oldCursor)
                {
                    AddDeletionAnnotation(
                        deletions,
                        NewOffsetForToken(newTokens, prefix + newCursor, newText.Length),
                        TextForTokenRange(oldText, oldTokens, prefix + oldCursor, match.Item1 - oldCursor));
                }
                MarkTokens(changed, newTokens, prefix + newCursor, match.Item2 - newCursor);
                oldCursor = match.Item1 + 1;
                newCursor = match.Item2 + 1;
            }

            if (oldCursor < oldMiddleCount)
            {
                AddDeletionAnnotation(
                    deletions,
                    NewOffsetForToken(newTokens, prefix + newCursor, newText.Length),
                    TextForTokenRange(oldText, oldTokens, prefix + oldCursor, oldMiddleCount - oldCursor));
            }
            MarkTokens(changed, newTokens, prefix + newCursor, newMiddleCount - newCursor);

            return new ParagraphDiff { ChangedMask = changed, Deletions = deletions };
        }

        private static int NewOffsetForToken(List<TextToken> tokens, int tokenIndex, int textLength)
        {
            return tokenIndex >= 0 && tokenIndex < tokens.Count ? tokens[tokenIndex].Start : textLength;
        }

        private static string TextForTokenRange(
            string text,
            List<TextToken> tokens,
            int start,
            int count)
        {
            if (count <= 0 || start < 0 || start >= tokens.Count)
            {
                return string.Empty;
            }
            int endIndex = Math.Min(tokens.Count - 1, start + count - 1);
            int textStart = tokens[start].Start;
            int textEnd = tokens[endIndex].Start + tokens[endIndex].Length;
            return text.Substring(textStart, textEnd - textStart);
        }

        private static void AddDeletionAnnotation(
            List<DeletionAnnotation> deletions,
            int offset,
            string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }
            deletions.Add(new DeletionAnnotation
            {
                NewTextOffset = offset,
                DeletedText = text
            });
        }

        private static List<TextToken> Tokenize(string text)
        {
            List<TextToken> tokens = new List<TextToken>();
            int index = 0;
            while (index < text.Length)
            {
                int start = index;
                char current = text[index];
                if (char.IsWhiteSpace(current))
                {
                    while (index < text.Length && char.IsWhiteSpace(text[index]))
                    {
                        index++;
                    }
                }
                else if (IsCjk(current))
                {
                    index++;
                }
                else if (char.IsLetterOrDigit(current) || current == '_')
                {
                    while (index < text.Length &&
                           !IsCjk(text[index]) &&
                           (char.IsLetterOrDigit(text[index]) || text[index] == '_'))
                    {
                        index++;
                    }
                }
                else
                {
                    index++;
                }

                tokens.Add(new TextToken
                {
                    Value = text.Substring(start, index - start),
                    Start = start,
                    Length = index - start
                });
            }
            return tokens;
        }

        private static bool IsCjk(char value)
        {
            int code = value;
            return (code >= 0x3400 && code <= 0x4DBF) ||
                   (code >= 0x4E00 && code <= 0x9FFF) ||
                   (code >= 0xF900 && code <= 0xFAFF) ||
                   (code >= 0x3040 && code <= 0x30FF) ||
                   (code >= 0xAC00 && code <= 0xD7AF);
        }

        private static void MarkTokens(bool[] changed, List<TextToken> tokens, int start, int count)
        {
            for (int index = start; index < start + count; index++)
            {
                TextToken token = tokens[index];
                for (int character = token.Start; character < token.Start + token.Length; character++)
                {
                    changed[character] = true;
                }
            }
        }

        private static XElement FindNeighborParagraph(
            List<ParagraphPair> alignment,
            List<ParagraphInfo> newParagraphs,
            int pairIndex,
            int direction)
        {
            int index = pairIndex + direction;
            while (index >= 0 && index < alignment.Count)
            {
                ParagraphPair candidate = alignment[index];
                if (candidate.NewIndex >= 0)
                {
                    return newParagraphs[candidate.NewIndex].Element;
                }
                index += direction;
            }
            return null;
        }

        private static void InsertDeletedParagraph(
            XDocument document,
            XElement previousParagraph,
            XElement nextParagraph,
            string deletedText)
        {
            XElement paragraph = new XElement(
                W + "p",
                new XElement(
                    W + "pPr",
                    new XElement(W + "spacing", new XAttribute(W + "before", "80"), new XAttribute(W + "after", "80"))),
                CreateDeletionMarkerRuns(deletedText));

            if (nextParagraph != null && nextParagraph.Parent != null)
            {
                nextParagraph.AddBeforeSelf(paragraph);
                return;
            }

            if (previousParagraph != null && previousParagraph.Parent != null)
            {
                XElement parent = previousParagraph.Parent;
                XElement sectionProperties = parent.Element(W + "sectPr");
                if (sectionProperties != null)
                {
                    sectionProperties.AddBeforeSelf(paragraph);
                }
                else
                {
                    parent.Add(paragraph);
                }
                return;
            }

            XElement body = document.Descendants(W + "body").FirstOrDefault();
            XElement container = body ?? document.Root;
            if (container == null)
            {
                return;
            }
            XElement finalSectionProperties = container.Element(W + "sectPr");
            if (finalSectionProperties != null)
            {
                finalSectionProperties.AddBeforeSelf(paragraph);
            }
            else
            {
                container.Add(paragraph);
            }
        }

        private static void AppendOrphanedDeletionSection(
            XDocument document,
            List<Tuple<string, string>> deletions)
        {
            XElement body = document.Descendants(W + "body").FirstOrDefault();
            if (body == null)
            {
                return;
            }

            XElement sectionProperties = body.Element(W + "sectPr");
            XElement heading = new XElement(
                W + "p",
                new XElement(
                    W + "pPr",
                    new XElement(W + "spacing", new XAttribute(W + "before", "240"), new XAttribute(W + "after", "100"))),
                CreateDeletionRun("旧版本中已删除的附加区域内容", false, true));
            AddBeforeSection(body, sectionProperties, heading);

            foreach (Tuple<string, string> deletion in deletions)
            {
                XElement sourceLabel = CreateDeletionRun("[来源：" + FriendlyPartName(deletion.Item1) + "] ", false, true);
                XElement paragraph = new XElement(
                    W + "p",
                    new XElement(
                        W + "pPr",
                        new XElement(W + "spacing", new XAttribute(W + "before", "60"), new XAttribute(W + "after", "60"))),
                    sourceLabel,
                    CreateDeletionMarkerRuns(deletion.Item2));
                AddBeforeSection(body, sectionProperties, paragraph);
            }
        }

        private static void AddBeforeSection(XElement body, XElement sectionProperties, XElement element)
        {
            if (sectionProperties != null)
            {
                sectionProperties.AddBeforeSelf(element);
            }
            else
            {
                body.Add(element);
            }
        }

        private static string FriendlyPartName(string partName)
        {
            if (partName.IndexOf("header", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "页眉";
            }
            if (partName.IndexOf("footer", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "页脚";
            }
            if (partName.IndexOf("footnotes", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "脚注";
            }
            if (partName.IndexOf("endnotes", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "尾注";
            }
            return "文档附加区域";
        }

        private static void InsertDeletionAnnotations(
            XElement paragraph,
            List<DeletionAnnotation> annotations)
        {
            if (annotations == null || annotations.Count == 0)
            {
                return;
            }

            List<DeletionAnnotation> grouped = annotations
                .GroupBy(annotation => annotation.NewTextOffset)
                .Select(group => new DeletionAnnotation
                {
                    NewTextOffset = group.Key,
                    DeletedText = string.Join("", group.Select(annotation => annotation.DeletedText))
                })
                .OrderByDescending(annotation => annotation.NewTextOffset)
                .ToList();

            foreach (DeletionAnnotation annotation in grouped)
            {
                InsertDeletionAnnotationAtOffset(paragraph, annotation.NewTextOffset, annotation.DeletedText);
            }
        }

        private static void InsertDeletionAnnotationAtOffset(
            XElement paragraph,
            int requestedOffset,
            string deletedText)
        {
            if (string.IsNullOrWhiteSpace(deletedText))
            {
                return;
            }

            int targetOffset = Math.Max(0, requestedOffset);
            List<XElement> markerRuns = CreateDeletionMarkerRuns(deletedText).ToList();
            List<XElement> runs = paragraph
                .Descendants(W + "r")
                .Where(run => run.Ancestors(W + "p").FirstOrDefault() == paragraph)
                .ToList();

            int offset = 0;
            foreach (XElement run in runs)
            {
                string runText = DirectRunText(run, paragraph);
                int runEnd = offset + runText.Length;
                if (targetOffset <= offset)
                {
                    run.AddBeforeSelf(markerRuns);
                    return;
                }
                if (targetOffset < runEnd)
                {
                    int splitAt = targetOffset - offset;
                    bool simpleTextRun = run.Elements()
                        .Where(element => element.Name != W + "rPr")
                        .All(element => element.Name == W + "t");
                    if (!simpleTextRun)
                    {
                        run.AddBeforeSelf(markerRuns);
                        return;
                    }

                    List<XElement> replacements = new List<XElement>();
                    if (splitAt > 0)
                    {
                        replacements.Add(CloneTextRun(run, runText.Substring(0, splitAt)));
                    }
                    replacements.AddRange(markerRuns);
                    if (splitAt < runText.Length)
                    {
                        replacements.Add(CloneTextRun(run, runText.Substring(splitAt)));
                    }
                    run.ReplaceWith(replacements);
                    return;
                }
                offset = runEnd;
            }

            XElement lastRun = runs.LastOrDefault();
            if (lastRun != null && lastRun.Parent != null)
            {
                lastRun.AddAfterSelf(markerRuns);
            }
            else
            {
                paragraph.Add(markerRuns);
            }
        }

        private static string DirectRunText(XElement run, XElement paragraph)
        {
            return string.Concat(
                run.Descendants(W + "t")
                    .Where(text => text.Ancestors(W + "r").FirstOrDefault() == run &&
                                   text.Ancestors(W + "p").FirstOrDefault() == paragraph)
                    .Select(text => text.Value));
        }

        private static XElement CloneTextRun(XElement originalRun, string text)
        {
            XElement run = new XElement(W + "r", originalRun.Attributes());
            XElement originalProperties = originalRun.Element(W + "rPr");
            if (originalProperties != null)
            {
                run.Add(new XElement(originalProperties));
            }
            XElement textElement = new XElement(W + "t", text);
            if (text.Length > 0 &&
                (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
            {
                textElement.SetAttributeValue(Xml + "space", "preserve");
            }
            run.Add(textElement);
            return run;
        }

        private static IEnumerable<XElement> CreateDeletionMarkerRuns(string deletedText)
        {
            yield return CreateDeletionRun("〔旧版删除：", false, true);
            yield return CreateDeletionRun(deletedText, true, false);
            yield return CreateDeletionRun("〕", false, true);
        }

        private static XElement CreateDeletionRun(string text, bool strike, bool bold)
        {
            XElement properties = new XElement(W + "rPr");
            if (bold)
            {
                properties.Add(new XElement(W + "b"));
            }
            if (strike)
            {
                properties.Add(new XElement(W + "strike"));
            }
            properties.Add(new XElement(W + "color", new XAttribute(W + "val", "C00000")));
            if (strike)
            {
                properties.Add(new XElement(W + "highlight", new XAttribute(W + "val", "lightGray")));
            }

            XElement textElement = new XElement(W + "t", text);
            if (text.Length > 0 &&
                (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
            {
                textElement.SetAttributeValue(Xml + "space", "preserve");
            }
            return new XElement(W + "r", properties, textElement);
        }

        private static void ApplyHighlight(XElement paragraph, bool[] changedMask, string highlightColor)
        {
            List<XElement> runs = paragraph
                .Descendants(W + "r")
                .Where(run => run.Ancestors(W + "p").FirstOrDefault() == paragraph)
                .ToList();
            int paragraphOffset = 0;

            foreach (XElement run in runs)
            {
                List<XElement> textNodes = run
                    .Descendants(W + "t")
                    .Where(text => text.Ancestors(W + "r").FirstOrDefault() == run &&
                                   text.Ancestors(W + "p").FirstOrDefault() == paragraph)
                    .ToList();
                string runText = string.Concat(textNodes.Select(text => text.Value));
                if (runText.Length == 0)
                {
                    continue;
                }

                int runStart = paragraphOffset;
                int runEnd = Math.Min(changedMask.Length, runStart + runText.Length);
                paragraphOffset += runText.Length;
                if (runStart >= changedMask.Length)
                {
                    continue;
                }

                bool anyChanged = false;
                bool anyUnchanged = false;
                for (int index = runStart; index < runEnd; index++)
                {
                    if (changedMask[index])
                    {
                        anyChanged = true;
                    }
                    else
                    {
                        anyUnchanged = true;
                    }
                }
                if (!anyChanged)
                {
                    continue;
                }

                if (!anyUnchanged)
                {
                    SetRunHighlight(run, highlightColor);
                    continue;
                }

                bool simpleTextRun = run.Elements()
                    .Where(element => element.Name != W + "rPr")
                    .All(element => element.Name == W + "t");
                if (!simpleTextRun)
                {
                    SetRunHighlight(run, highlightColor);
                    continue;
                }

                List<XElement> splitRuns = new List<XElement>();
                int localStart = 0;
                while (localStart < runText.Length)
                {
                    bool isChanged = changedMask[runStart + localStart];
                    int localEnd = localStart + 1;
                    while (localEnd < runText.Length &&
                           changedMask[runStart + localEnd] == isChanged)
                    {
                        localEnd++;
                    }

                    string segmentText = runText.Substring(localStart, localEnd - localStart);
                    XElement splitRun = CreateSplitRun(run, segmentText, isChanged, highlightColor);
                    splitRuns.Add(splitRun);
                    localStart = localEnd;
                }

                run.ReplaceWith(splitRuns);
            }
        }

        private static XElement CreateSplitRun(
            XElement originalRun,
            string text,
            bool highlighted,
            string highlightColor)
        {
            XElement run = new XElement(W + "r", originalRun.Attributes());
            XElement originalProperties = originalRun.Element(W + "rPr");
            if (originalProperties != null)
            {
                run.Add(new XElement(originalProperties));
            }

            XElement textElement = new XElement(W + "t", text);
            if (text.Length > 0 &&
                (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
            {
                textElement.SetAttributeValue(Xml + "space", "preserve");
            }
            run.Add(textElement);
            if (highlighted)
            {
                SetRunHighlight(run, highlightColor);
            }
            return run;
        }

        private static void SetRunHighlight(XElement run, string highlightColor)
        {
            XElement properties = run.Element(W + "rPr");
            if (properties == null)
            {
                properties = new XElement(W + "rPr");
                run.AddFirst(properties);
            }

            foreach (XElement existing in properties.Elements(W + "highlight").ToList())
            {
                existing.Remove();
            }
            XElement highlight = new XElement(W + "highlight", new XAttribute(W + "val", highlightColor));
            HashSet<XName> elementsAfterHighlight = new HashSet<XName>
            {
                W + "u", W + "effect", W + "bdr", W + "shd", W + "fitText",
                W + "vertAlign", W + "rtl", W + "cs", W + "em", W + "lang",
                W + "eastAsianLayout", W + "specVanish", W + "oMath", W + "rPrChange"
            };
            XElement insertionPoint = properties.Elements().FirstOrDefault(element => elementsAfterHighlight.Contains(element.Name));
            if (insertionPoint == null)
            {
                properties.Add(highlight);
            }
            else
            {
                insertionPoint.AddBeforeSelf(highlight);
            }
        }

        private static int CountTrue(bool[] values)
        {
            int count = 0;
            foreach (bool value in values)
            {
                if (value)
                {
                    count++;
                }
            }
            return count;
        }

        private static XDocument LoadXmlSafely(byte[] bytes, string partName)
        {
            try
            {
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.DtdProcessing = DtdProcessing.Prohibit;
                settings.XmlResolver = null;
                settings.MaxCharactersInDocument = 64L * 1024L * 1024L;

                using (MemoryStream stream = new MemoryStream(bytes, false))
                using (XmlReader reader = XmlReader.Create(stream, settings))
                {
                    return XDocument.Load(reader, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);
                }
            }
            catch (Exception exception)
            {
                throw new InvalidDataException("无法读取 DOCX 内部内容：" + partName, exception);
            }
        }

        private static byte[] SaveXml(XDocument document)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                settings.Indent = false;
                settings.OmitXmlDeclaration = document.Declaration == null;
                using (XmlWriter writer = XmlWriter.Create(stream, settings))
                {
                    document.Save(writer);
                }
                return stream.ToArray();
            }
        }

        private static byte[] ReadEntryBytes(ZipArchiveEntry entry, int maximumBytes)
        {
            if (entry.Length > maximumBytes)
            {
                throw new InvalidDataException("DOCX 内部 XML 过大，已停止处理：" + entry.FullName);
            }

            using (Stream source = entry.Open())
            using (MemoryStream destination = new MemoryStream())
            {
                byte[] buffer = new byte[81920];
                int total = 0;
                int read;
                while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += read;
                    if (total > maximumBytes)
                    {
                        throw new InvalidDataException("DOCX 内部 XML 解压后过大，已停止处理：" + entry.FullName);
                    }
                    destination.Write(buffer, 0, read);
                }
                return destination.ToArray();
            }
        }

        private static void CopyEntry(
            ZipArchiveEntry sourceEntry,
            ZipArchive outputArchive,
            Dictionary<string, byte[]> modifiedParts)
        {
            CompressionLevel compression = IsAlreadyCompressed(sourceEntry.FullName)
                ? CompressionLevel.NoCompression
                : CompressionLevel.Optimal;
            ZipArchiveEntry targetEntry = outputArchive.CreateEntry(sourceEntry.FullName, compression);
            targetEntry.LastWriteTime = sourceEntry.LastWriteTime;

            byte[] replacement;
            string partName = NormalizePartName(sourceEntry.FullName);
            using (Stream target = targetEntry.Open())
            {
                if (modifiedParts.TryGetValue(partName, out replacement))
                {
                    target.Write(replacement, 0, replacement.Length);
                }
                else
                {
                    using (Stream source = sourceEntry.Open())
                    {
                        source.CopyTo(target);
                    }
                }
            }
        }

        private static bool IsAlreadyCompressed(string name)
        {
            string extension = Path.GetExtension(name).ToLowerInvariant();
            return extension == ".png" || extension == ".jpg" || extension == ".jpeg" ||
                   extension == ".gif" || extension == ".zip" || extension == ".mp3" ||
                   extension == ".mp4";
        }

        private static void Report(IProgress<int> progress, int value)
        {
            if (progress != null)
            {
                progress.Report(Math.Max(0, Math.Min(100, value)));
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }
    }
}

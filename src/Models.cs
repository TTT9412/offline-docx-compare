using System;

namespace OfflineDocxCompare
{
    public sealed class CompareResult
    {
        public string OutputPath { get; set; }
        public int ComparedParts { get; set; }
        public int ChangedParagraphs { get; set; }
        public int InsertedParagraphs { get; set; }
        public int ReplacedParagraphs { get; set; }
        public int DeletedParagraphs { get; set; }
        public int HighlightedCharacters { get; set; }
        public int DeletedCharactersShown { get; set; }

        public bool HasDifferences
        {
            get
            {
                return ChangedParagraphs > 0 || DeletedParagraphs > 0;
            }
        }
    }

    internal sealed class ParagraphPair
    {
        public ParagraphPair(int oldIndex, int newIndex)
        {
            OldIndex = oldIndex;
            NewIndex = newIndex;
        }

        public int OldIndex { get; private set; }
        public int NewIndex { get; private set; }
    }

    internal sealed class TextToken
    {
        public string Value { get; set; }
        public int Start { get; set; }
        public int Length { get; set; }
    }

    internal sealed class ParagraphInfo
    {
        public System.Xml.Linq.XElement Element { get; set; }
        public string Text { get; set; }
        public System.Collections.Generic.HashSet<int> Fingerprint { get; set; }
    }

    internal sealed class DeletionAnnotation
    {
        public int NewTextOffset { get; set; }
        public string DeletedText { get; set; }
    }

    internal sealed class ParagraphDiff
    {
        public bool[] ChangedMask { get; set; }
        public System.Collections.Generic.List<DeletionAnnotation> Deletions { get; set; }
    }
}

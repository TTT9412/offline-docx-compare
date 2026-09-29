using OfflineDocxCompare;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

internal static class StressTestRunner
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly string SymbolText = "© ® ™ € ¥ £ ± × ÷ ≠ ≤ ≥ → ← ↔ ∑ ∫ √ ∞ § ¶ • … — ‘引号’ \"quotes\" 😀 🚀";

    private static int Main(string[] args)
    {
        string root = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "OfflineDocxCompareStress");
        Directory.CreateDirectory(root);
        try
        {
            List<string> results = new List<string>();
            results.Add(RunMixedContent(root));
            results.Add(RunLargeDocument(root));
            results.Add(RunDeletionHeavy(root));
            results.Add(RunFormattingNoise(root));
            Console.WriteLine("STRESS PASS: 4 rounds");
            foreach (string result in results) Console.WriteLine(result);
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine("STRESS FAIL: " + exception);
            return 1;
        }
    }

    private static string RunMixedContent(string root)
    {
        string oldPath = Path.Combine(root, "01-mixed-old.docx");
        string newPath = Path.Combine(root, "01-mixed-new.docx");
        string outputPath = Path.Combine(root, "01-mixed-output.docx");
        string oldBody = BuildParagraphs(700, false, false, 0) + BuildTable(45, 7, false) + BuildTextbox(false);
        string newBody = BuildParagraphs(700, true, false, 0) + BuildTable(45, 7, true) + BuildTextbox(true);
        CreatePackage(oldPath, oldBody, "机密页眉 / CONFIDENTIAL / Rev-A", "第 1 页 · Offline", "脚注：旧税率为 6%。", "尾注：Legacy appendix α.");
        CreatePackage(newPath, newBody, "机密页眉 / CONFIDENTIAL / Rev-B", "第 1 页 · Local only", "脚注：新税率为 9%。", "尾注：Revised appendix β.");
        CompareResult result;
        long elapsed = CompareTimed(oldPath, newPath, outputPath, "yellow", out result);
        ValidatePackage(outputPath);
        Assert(result.ChangedParagraphs > 25, "混合内容变化计数过少");
        Assert(result.HighlightedCharacters > 100, "混合内容高亮字符过少");
        Assert(result.DeletedCharactersShown > 50, "混合内容删除展示过少");
        Assert(PartHasHighlight(outputPath, "word/document.xml", "yellow"), "正文缺少黄色高亮");
        Assert(PartHasHighlight(outputPath, "word/header1.xml", "yellow"), "页眉缺少黄色高亮");
        Assert(ReadAllText(outputPath).Contains("旧版删除："), "没有显示旧版删除内容");
        return FormatResult("混合语言/符号/全部文字部件", oldPath, result, elapsed);
    }

    private static string RunLargeDocument(string root)
    {
        string oldPath = Path.Combine(root, "02-large-old.docx");
        string newPath = Path.Combine(root, "02-large-new.docx");
        string outputPath = Path.Combine(root, "02-large-output.docx");
        string oldLong = Repeat("旧版超长段落包含中文Alpha-2026与符号→∞；", 460);
        string newLong = Repeat("新版超长段落包含中文Beta-2027与符号↔√；", 470);
        string oldBody = BuildParagraphs(1800, false, false, 10000) + Paragraph(oldLong, 1, "111111");
        string newBody = BuildParagraphs(1800, true, false, 10000) + Paragraph(newLong, 5, "222222");
        CreatePackage(oldPath, oldBody, "大型文档 A", "Footer A", "Large footnote A", "Large endnote A");
        CreatePackage(newPath, newBody, "大型文档 B", "Footer B", "Large footnote B", "Large endnote B");
        CompareResult result;
        long elapsed = CompareTimed(oldPath, newPath, outputPath, "blue", out result);
        ValidatePackage(outputPath);
        Assert(new FileInfo(outputPath).Length > 30000, "大型输出文件异常偏小");
        Assert(result.ChangedParagraphs > 40, "大型文档变化计数过少");
        Assert(PartHasHighlight(outputPath, "word/document.xml", "blue"), "大型文档缺少蓝色高亮");
        Assert(elapsed < 120000, "大型文档耗时超过 120 秒");
        return FormatResult("1800 段及超长段落", oldPath, result, elapsed);
    }

    private static string RunDeletionHeavy(string root)
    {
        string oldPath = Path.Combine(root, "03-deletion-old.docx");
        string newPath = Path.Combine(root, "03-deletion-new.docx");
        string outputPath = Path.Combine(root, "03-deletion-output.docx");
        StringBuilder oldBody = new StringBuilder();
        StringBuilder newBody = new StringBuilder();
        for (int i = 0; i < 620; i++)
        {
            string text = BaseText(i + 20000);
            oldBody.Append(Paragraph(text + (i % 11 == 0 ? " 旧版独有段落 DELETE-ME-" + i : ""), 3, "444444"));
            if (i % 7 != 0) newBody.Append(Paragraph(text, 4, "006666"));
            if (i % 29 == 0) newBody.Append(Paragraph("新增审计说明 INSERTED-" + i + "：仅在新版本出现。", 2, "008888"));
        }
        CreatePackage(oldPath, oldBody.ToString(), "将被整体删除的页眉", "旧页脚", "被删除脚注", "被删除尾注");
        CreatePackage(newPath, newBody.ToString(), "保留页眉", "新页脚", "新脚注", "新尾注");
        CompareResult result;
        long elapsed = CompareTimed(oldPath, newPath, outputPath, "green", out result);
        ValidatePackage(outputPath);
        XDocument document = ReadPart(outputPath, "word/document.xml");
        Assert(result.DeletedParagraphs >= 70, "大量整段删除未被完整识别");
        Assert(document.Descendants(W + "strike").Any(), "删除内容没有删除线");
        Assert(ReadAllText(outputPath).Contains("DELETE-ME-"), "旧版独有内容未写回副本");
        Assert(PartHasHighlight(outputPath, "word/document.xml", "green"), "新增内容缺少绿色高亮");
        return FormatResult("大量整段删除及新增", oldPath, result, elapsed);
    }

    private static string RunFormattingNoise(string root)
    {
        string oldPath = Path.Combine(root, "04-format-old.docx");
        string newPath = Path.Combine(root, "04-format-new.docx");
        string outputPath = Path.Combine(root, "04-format-output.docx");
        string oldBody = BuildParagraphs(950, false, true, 30000);
        string newBody = BuildParagraphs(950, true, true, 30000);
        CreatePackage(oldPath, oldBody, "相同页眉", "相同页脚", "相同脚注", "相同尾注");
        CreatePackage(newPath, newBody, "相同页眉", "相同页脚", "相同脚注", "相同尾注");
        CompareResult result;
        long elapsed = CompareTimed(oldPath, newPath, outputPath, "cyan", out result);
        ValidatePackage(outputPath);
        Assert(!result.HasDifferences, "纯格式和分段变化被误报为文字差异");
        Assert(result.HighlightedCharacters == 0 && result.DeletedCharactersShown == 0, "纯格式场景产生了标注");
        Assert(!ReadPart(outputPath, "word/document.xml").Descendants(W + "highlight").Any(), "纯格式场景写入了高亮");
        Assert(HashPart(newPath, "word/document.xml") == HashPart(outputPath, "word/document.xml"), "无文字差异时正文不应被改写");
        return FormatResult("950 段纯格式噪声", oldPath, result, elapsed);
    }

    private static string BuildParagraphs(int count, bool newer, bool formattingOnly, int idOffset)
    {
        StringBuilder body = new StringBuilder();
        for (int i = 0; i < count; i++)
        {
            string text = BaseText(i + idOffset);
            if (!formattingOnly)
            {
                if (i % 37 == 0) text += newer ? " | STATUS=BETA; 金额=€12,345.67" : " | STATUS=ALPHA; 金额=€10,000.00";
                if (i % 53 == 0 && newer) text += " | 新增条款：数据必须留在本机。";
                if (i % 61 == 0 && !newer) text += " | 旧版待删除短语 LEGACY-ONLY.";
            }
            body.Append(Paragraph(text, newer ? ((i % 5) + 1) : 1, newer ? "006666" : "990000"));
        }
        return body.ToString();
    }

    private static string BaseText(int id)
    {
        return "第" + id.ToString("D5") + "节：本地合同审阅 Local contract review; customer=Acme_" + id +
            "; formula=(a+b)^2=a²+2ab+b²; path=C:\\Archive\\Q" + (id % 4 + 1) +
            "; email=test+" + id + "@example.invalid; symbols=" + SymbolText + ".";
    }

    private static string Paragraph(string text, int fragments, string color)
    {
        StringBuilder xml = new StringBuilder("<w:p>");
        int size = Math.Max(1, (text.Length + fragments - 1) / fragments);
        for (int start = 0; start < text.Length; start += size)
        {
            string piece = text.Substring(start, Math.Min(size, text.Length - start));
            xml.Append("<w:r><w:rPr><w:color w:val=\"").Append(color).Append("\"/>");
            if ((start / size) % 2 == 1) xml.Append("<w:b/>");
            xml.Append("</w:rPr><w:t xml:space=\"preserve\">").Append(Escape(piece)).Append("</w:t></w:r>");
        }
        return xml.Append("</w:p>").ToString();
    }

    private static string BuildTable(int rows, int columns, bool newer)
    {
        StringBuilder xml = new StringBuilder("<w:tbl><w:tblPr><w:tblBorders><w:top w:val=\"single\"/><w:left w:val=\"single\"/><w:bottom w:val=\"single\"/><w:right w:val=\"single\"/></w:tblBorders></w:tblPr>");
        for (int r = 0; r < rows; r++)
        {
            xml.Append("<w:tr>");
            for (int c = 0; c < columns; c++)
            {
                string text = "表格 R" + r + "C" + c + " / SKU-" + (r * columns + c).ToString("D4") + " / 数量=" + (r + c + 1);
                if (newer && r % 13 == 0 && c == 3) text += " / 已复核 ✓";
                xml.Append("<w:tc>").Append(Paragraph(text, newer ? 2 : 1, newer ? "005555" : "333333"));
                if (r == 2 && c == 2)
                {
                    xml.Append("<w:tbl><w:tr><w:tc>").Append(Paragraph(newer ? "嵌套表格 NEW Ω" : "嵌套表格 OLD α", 2, "555555")).Append("</w:tc></w:tr></w:tbl>");
                }
                xml.Append("</w:tc>");
            }
            xml.Append("</w:tr>");
        }
        return xml.Append("</w:tbl>").ToString();
    }

    private static string BuildTextbox(bool newer)
    {
        return "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
            Paragraph(newer ? "文本框 TextBox: NEW value ② → 已确认" : "文本框 TextBox: OLD value ① → 待确认", 2, "333333") +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";
    }

    private static void CreatePackage(string path, string body, string header, string footer, string footnote, string endnote)
    {
        if (File.Exists(path)) File.Delete(path);
        using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false))
        {
            Write(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/><Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/><Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/><Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/></Types>");
            Write(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Write(archive, "word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/></Relationships>");
            Write(archive, "word/document.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:v=\"urn:schemas-microsoft-com:vml\"><w:body>" + body + "<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r><w:r><w:endnoteReference w:id=\"1\"/></w:r></w:p><w:sectPr><w:headerReference w:type=\"default\" r:id=\"rId1\"/><w:footerReference w:type=\"default\" r:id=\"rId2\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr></w:body></w:document>");
            Write(archive, "word/header1.xml", Story("hdr", header));
            Write(archive, "word/footer1.xml", Story("ftr", footer));
            Write(archive, "word/footnotes.xml", Notes("footnotes", "footnote", footnote));
            Write(archive, "word/endnotes.xml", Notes("endnotes", "endnote", endnote));
        }
    }

    private static string Story(string root, string text)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:" + root + " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + Paragraph(text, 2, "333333") + "</w:" + root + ">";
    }

    private static string Notes(string root, string child, string text)
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?><w:" + root + " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:" + child + " w:id=\"-1\"><w:p/></w:" + child + "><w:" + child + " w:id=\"0\"><w:p/></w:" + child + "><w:" + child + " w:id=\"1\">" + Paragraph(text, 2, "333333") + "</w:" + child + "></w:" + root + ">";
    }

    private static long CompareTimed(string oldPath, string newPath, string outputPath, string color, out CompareResult result)
    {
        Stopwatch timer = Stopwatch.StartNew();
        result = DocumentComparer.Compare(oldPath, newPath, outputPath, color, null);
        timer.Stop();
        return timer.ElapsedMilliseconds;
    }

    private static void ValidatePackage(string path)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            Assert(archive.GetEntry("[Content_Types].xml") != null, "DOCX 缺少内容类型声明");
            foreach (ZipArchiveEntry entry in archive.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                using (Stream stream = entry.Open())
                using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024L * 1024L }))
                {
                    while (reader.Read()) { }
                }
            }
        }
    }

    private static XDocument ReadPart(string path, string part)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        using (Stream stream = archive.GetEntry(part).Open()) return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string ReadAllText(string path)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            StringBuilder text = new StringBuilder();
            foreach (ZipArchiveEntry entry in archive.Entries.Where(e => e.FullName.StartsWith("word/") && e.FullName.EndsWith(".xml")))
            {
                using (Stream stream = entry.Open())
                {
                    XDocument document = XDocument.Load(stream);
                    foreach (XElement node in document.Descendants(W + "t")) text.Append(node.Value);
                }
            }
            return text.ToString();
        }
    }

    private static bool PartHasHighlight(string path, string part, string color)
    {
        return ReadPart(path, part).Descendants(W + "highlight").Any(e => string.Equals((string)e.Attribute(W + "val"), color, StringComparison.OrdinalIgnoreCase));
    }

    private static string HashPart(string path, string part)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        using (Stream stream = archive.GetEntry(part).Open())
        using (SHA256 sha = SHA256.Create()) return Convert.ToBase64String(sha.ComputeHash(stream));
    }

    private static string FormatResult(string name, string input, CompareResult result, long elapsed)
    {
        long chars = ReadAllText(input).Length;
        return "[PASS] " + name + " | chars=" + chars + " | changed=" + result.ChangedParagraphs + " | deleted=" + result.DeletedParagraphs + " | ms=" + elapsed;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using (Stream stream = entry.Open())
        using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false))) writer.Write(content);
    }

    private static string Repeat(string value, int count)
    {
        StringBuilder result = new StringBuilder(value.Length * count);
        for (int i = 0; i < count; i++) result.Append(value);
        return result.ToString();
    }

    private static string Escape(string value)
    {
        return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}

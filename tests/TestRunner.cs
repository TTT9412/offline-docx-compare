using OfflineDocxCompare;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

internal static class TestRunner
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static int passed;

    private static int Main(string[] args)
    {
        string directory = args.Length > 0 ? args[0] : Path.Combine(Path.GetTempPath(), "OfflineDocxCompareTests");
        Directory.CreateDirectory(directory);

        try
        {
            Run("正文、表格和页眉的文字变化会被高亮", delegate { TestBodyTableAndHeader(directory); });
            Run("仅格式变化不会被识别为文字差异", delegate { TestFormattingOnly(directory); });
            Run("纯删除段落会以删除线写回新副本", delegate { TestDeletionOnly(directory); });
            Run("整个被删除的页眉区域会汇总到副本", delegate { TestRemovedHeaderPart(directory); });
            Run("输出路径不能覆盖源文件", delegate { TestSourceOverwriteProtection(directory); });
            Console.WriteLine("PASS: " + passed + " tests");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine("FAIL: " + exception);
            return 1;
        }
    }

    private static void Run(string name, Action test)
    {
        test();
        passed++;
        Console.WriteLine("[PASS] " + name);
    }

    private static void TestBodyTableAndHeader(string directory)
    {
        string oldPath = Path.Combine(directory, "old-main.docx");
        string newPath = Path.Combine(directory, "new-main.docx");
        string outputPath = Path.Combine(directory, "output-main.docx");

        CreateDocx(
            oldPath,
            "<w:p><w:r><w:t>合同金额为</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>100元</w:t></w:r></w:p>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>联系人：张三</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>文本框旧内容</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>",
            "版本A");
        CreateDocx(
            newPath,
            "<w:p><w:r><w:t>合同金额为</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>120元</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>新增条款</w:t></w:r></w:p>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>联系人：李四</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>文本框新内容</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>",
            "版本B");

        CompareResult result = DocumentComparer.Compare(oldPath, newPath, outputPath, "yellow", null);
        Assert(File.Exists(outputPath), "没有生成输出文件");
        Assert(result.ChangedParagraphs >= 3, "正文、表格或页眉变化未全部计入");
        Assert(result.HighlightedCharacters > 0, "没有高亮任何字符");
        Assert(result.DeletedCharactersShown > 0, "被替换的旧文字没有计入删除展示统计");

        XDocument document = ReadXmlPart(outputPath, "word/document.xml");
        XDocument header = ReadXmlPart(outputPath, "word/header1.xml");
        Assert(HasHighlight(document, "yellow"), "正文中没有黄色高亮");
        Assert(HasHighlight(header, "yellow"), "页眉中没有黄色高亮");
        string outputText = string.Concat(document.Descendants(W + "t").Select(node => node.Value));
        Assert(outputText.Contains("120元"), "新金额丢失");
        Assert(outputText.Contains("旧版删除："), "替换位置没有显示旧版删除标记");
        Assert(outputText.Contains("100"), "被替换的旧金额没有体现在副本中");
        Assert(document.Descendants(W + "r")
            .Any(run => run.Descendants(W + "t").Any(node => node.Value.Contains("100")) &&
                        run.Descendants(W + "strike").Any()), "被替换的旧金额没有使用删除线");
        Assert(outputText.Contains("文本框") && outputText.Contains("新内容"), "文本框新版本文字丢失");
        Assert(document.Descendants(W + "p")
            .Where(paragraph => paragraph.Ancestors(W + "txbxContent").Any())
            .SelectMany(paragraph => paragraph.Descendants(W + "highlight"))
            .Any(), "文本框文字变化没有高亮");

        XElement amountRun = document.Descendants(W + "r")
            .FirstOrDefault(run => string.Concat(run.Descendants(W + "t").Select(node => node.Value)).Contains("120"));
        Assert(amountRun != null && amountRun.Element(W + "rPr") != null && amountRun.Element(W + "rPr").Element(W + "b") != null,
            "新版本原有的粗体格式未保留");

        Assert(ZipEntryCount(outputPath) == ZipEntryCount(newPath), "输出 DOCX 包结构条目数量发生意外变化");
    }

    private static void TestFormattingOnly(string directory)
    {
        string oldPath = Path.Combine(directory, "old-format.docx");
        string newPath = Path.Combine(directory, "new-format.docx");
        string outputPath = Path.Combine(directory, "output-format.docx");
        CreateDocx(
            oldPath,
            "<w:p><w:r><w:rPr><w:color w:val=\"FF0000\"/></w:rPr><w:t>文字完全相同</w:t></w:r></w:p>",
            "相同页眉");
        CreateDocx(
            newPath,
            "<w:p><w:r><w:rPr><w:color w:val=\"0000FF\"/></w:rPr><w:t>文字完全相同</w:t></w:r></w:p>",
            "相同页眉");

        CompareResult result = DocumentComparer.Compare(oldPath, newPath, outputPath, "blue", null);
        XDocument document = ReadXmlPart(outputPath, "word/document.xml");
        Assert(result.ChangedParagraphs == 0, "仅文字颜色变化被错误识别为内容变化");
        Assert(!document.Descendants(W + "highlight").Any(), "仅格式变化时不应写入高亮");
        Assert(document.Descendants(W + "color").Any(node => (string)node.Attribute(W + "val") == "0000FF"),
            "新版本文字颜色未保留");
    }

    private static void TestDeletionOnly(string directory)
    {
        string oldPath = Path.Combine(directory, "old-delete.docx");
        string newPath = Path.Combine(directory, "new-delete.docx");
        string outputPath = Path.Combine(directory, "output-delete.docx");
        CreateDocx(
            oldPath,
            "<w:p><w:r><w:t>保留内容</w:t></w:r></w:p><w:p><w:r><w:t>将被删除</w:t></w:r></w:p>",
            "页眉");
        CreateDocx(
            newPath,
            "<w:p><w:r><w:t>保留内容</w:t></w:r></w:p>",
            "页眉");

        CompareResult result = DocumentComparer.Compare(oldPath, newPath, outputPath, "green", null);
        XDocument document = ReadXmlPart(outputPath, "word/document.xml");
        Assert(result.DeletedParagraphs == 1, "纯删除段落统计不正确");
        Assert(result.DeletedCharactersShown == "将被删除".Length, "纯删除字符统计不正确");
        Assert(result.ChangedParagraphs == 0, "纯删除不应制造新版本文字高亮");
        string outputText = string.Concat(document.Descendants(W + "t").Select(node => node.Value));
        Assert(outputText.Contains("将被删除"), "纯删除段落没有插回副本");
        Assert(outputText.Contains("旧版删除："), "纯删除段落缺少删除标记");
        Assert(document.Descendants(W + "r")
            .Any(run => run.Descendants(W + "t").Any(node => node.Value.Contains("将被删除")) &&
                        run.Descendants(W + "strike").Any()), "纯删除段落没有使用红色删除线");
    }

    private static void TestSourceOverwriteProtection(string directory)
    {
        string oldPath = Path.Combine(directory, "old-protect.docx");
        string newPath = Path.Combine(directory, "new-protect.docx");
        CreateDocx(oldPath, "<w:p><w:r><w:t>A</w:t></w:r></w:p>", "H");
        CreateDocx(newPath, "<w:p><w:r><w:t>B</w:t></w:r></w:p>", "H");

        bool threw = false;
        try
        {
            DocumentComparer.Compare(oldPath, newPath, newPath, "yellow", null);
        }
        catch (ArgumentException)
        {
            threw = true;
        }
        Assert(threw, "程序允许输出覆盖源文件");
    }

    private static void TestRemovedHeaderPart(string directory)
    {
        string oldPath = Path.Combine(directory, "old-header.docx");
        string newPath = Path.Combine(directory, "new-header.docx");
        string outputPath = Path.Combine(directory, "output-header.docx");
        CreateDocx(oldPath, "<w:p><w:r><w:t>正文</w:t></w:r></w:p>", "已删除的页眉");
        CreateDocx(newPath, "<w:p><w:r><w:t>正文</w:t></w:r></w:p>", "", false);

        CompareResult result = DocumentComparer.Compare(oldPath, newPath, outputPath, "yellow", null);
        XDocument document = ReadXmlPart(outputPath, "word/document.xml");
        string outputText = string.Concat(document.Descendants(W + "t").Select(node => node.Value));
        Assert(result.DeletedParagraphs == 1, "被删除页眉未计入删除统计");
        Assert(outputText.Contains("已删除的页眉"), "被删除页眉没有汇总到副本正文");
        Assert(outputText.Contains("来源：页眉"), "被删除页眉缺少来源说明");
        Assert(document.Descendants(W + "r")
            .Any(run => run.Descendants(W + "t").Any(node => node.Value.Contains("已删除的页眉")) &&
                        run.Descendants(W + "strike").Any()), "被删除页眉没有使用删除线");
    }

    private static void CreateDocx(string path, string bodyContent, string headerText, bool includeHeader = true)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }

        using (FileStream stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false))
        {
            WriteEntry(archive, "[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
                (includeHeader ? "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" : "") +
                "</Types>");
            WriteEntry(archive, "_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
                "</Relationships>");
            WriteEntry(archive, "word/_rels/document.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                (includeHeader ? "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" : "") +
                "</Relationships>");
            WriteEntry(archive, "word/document.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
                "<w:body>" + bodyContent +
                "<w:sectPr>" + (includeHeader ? "<w:headerReference w:type=\"default\" r:id=\"rId1\"/>" : "") + "<w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>" +
                "</w:body></w:document>");
            if (includeHeader)
            {
                WriteEntry(archive, "word/header1.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:p><w:r><w:t>" +
                    EscapeXml(headerText) +
                    "</w:t></w:r></w:p></w:hdr>");
            }
        }
    }

    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using (Stream stream = entry.Open())
        using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false)))
        {
            writer.Write(content);
        }
    }

    private static XDocument ReadXmlPart(string path, string partName)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            ZipArchiveEntry entry = archive.GetEntry(partName);
            Assert(entry != null, "缺少 DOCX 部件：" + partName);
            using (Stream stream = entry.Open())
            {
                return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }
        }
    }

    private static bool HasHighlight(XDocument document, string color)
    {
        return document.Descendants(W + "highlight")
            .Any(node => string.Equals((string)node.Attribute(W + "val"), color, StringComparison.OrdinalIgnoreCase));
    }

    private static int ZipEntryCount(string path)
    {
        using (ZipArchive archive = ZipFile.OpenRead(path))
        {
            return archive.Entries.Count;
        }
    }

    private static string EscapeXml(string value)
    {
        return value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}

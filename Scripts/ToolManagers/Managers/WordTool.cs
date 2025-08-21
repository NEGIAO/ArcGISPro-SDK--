using Aspose.Words.Replacing;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CCTool.Scripts.ToolManagers.Managers
{
    public class WordTool
    {
        // 打开Document
        public static Document OpenDocument(string wordPath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Aspose.Words.LoadOptions loadOptinos = new Aspose.Words.LoadOptions(Aspose.Words.LoadFormat.Auto, null, null);

            Document doc = new Document(wordPath);
            return doc;
        }

        // 替换Word中的文本
        public static void WordRepalceText(string wordPath, string targetText, string replaceText)
        {
            // 打开Word
            Document doc = OpenDocument(wordPath);
            // 执行替换操作
            doc.Range.Replace(targetText, replaceText, new FindReplaceOptions(FindReplaceDirection.Forward));
            // 保存修改后的文档
            doc.Save(wordPath);
        }

        // word插入图片
        public static void WordInsertPic(string wordPath, string imagePath)
        {
            // 打开Word
            Document doc = OpenDocument(wordPath);
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            // 插入图片
            builder.InsertImage(imagePath);

            // 保存修改后的文档
            doc.Save(wordPath);
        }

        // word插入图片
        public static void WordInsertPic(string wordPath, List<string> imagePaths)
        {
            // 打开Word
            Document doc = OpenDocument(wordPath);
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();

            // 插入图片
            foreach (string imagePath in imagePaths)
            {
                builder.InsertImage(imagePath);
            }

            // 保存修改后的文档
            doc.Save(wordPath);
        }

        // HTML导入Word
        public static void Html2Word(string htmlPath, string wordPath)
        {
            // 打开Word
            Document doc = OpenDocument(wordPath);

            // 加载 HTML 文件
            Document htmlDoc = new Document(htmlPath);

            // 创建 DocumentBuilder 用于操作目标 Word 文档
            DocumentBuilder builder = new DocumentBuilder(doc);

            // 将光标移动到文档的末尾（或其他需要插入的地方）
            builder.MoveToDocumentEnd();

            // 将 HTML 文档的内容插入到 Word 文档中
            foreach (Section section in htmlDoc.Sections)
            {
                // 导入 HTML 中的每个部分，并插入到目标文档中
                Section importedSection = (Section)doc.ImportNode(section, true);
                doc.AppendChild(importedSection);
            }

            // 保存修改后的 Word 文档
            doc.Save(wordPath);
        }

        public static void ImportToPDF(string wordPath, string pdfPath) 
        {
            // 加载 Word 文档
            Document doc = new Document(wordPath);
            // 保存为 PDF
            doc.Save(pdfPath, SaveFormat.Pdf);

        }
    }
}

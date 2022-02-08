namespace PflanzenPaesse.Repositories.WordRepository
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public static class WordRepository
    {
        public static readonly string[] AllowedFileEndings = {"docx"};

        public static Body Import(string fileName)
        {
            Console.WriteLine("Importing template...");
            using var wordDocument = WordprocessingDocument.Open(fileName, false);
            var mainPart = wordDocument.MainDocumentPart;
            var document = mainPart.Document;
            var body = document.Body;
            return body;
        }

        public static Body Replace(Body body, IDictionary<string, string> mapping)
        {
            Console.WriteLine("Inserting values...");
            foreach(var map in mapping)
            {
                var text = body.OuterXml;
                text = text.Replace($"${{{map.Key}}}", map.Value);
                body = new Body(text);
            }

            return body;
        }

        public static async Task ExportAsync(string filename, IEnumerable<Body> bodies)
        {
            Console.WriteLine("Exporting document...");
            using var wordDocument = WordprocessingDocument.Create(filename, WordprocessingDocumentType.Document);
            var mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());
            foreach(var insertBody in bodies)
            {
                foreach(var child in insertBody.ChildElements.OfType<Paragraph>())
                {
                    body.AppendChild(new Paragraph(child.OuterXml));
                }
            }
            using var writer = new StreamWriter(wordDocument.MainDocumentPart.GetStream(FileMode.Create));
            await writer.FlushAsync();
        }
    }
}

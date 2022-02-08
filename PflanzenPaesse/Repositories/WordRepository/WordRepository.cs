namespace PflanzenPaesse.Repositories.WordRepository
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public static class WordRepository
    {
        public static readonly string[] AllowedFileEndings = {"docx"};

        public static async Task BuildPaesseAsync(string templateFileName, IEnumerable<IDictionary<string, string>> mapping, string outputFileName)
        {
            Console.WriteLine("Importing template...");
            using var templateWordDocument = WordprocessingDocument.Open(templateFileName, false);
            var templateMainPart = templateWordDocument.MainDocumentPart;
            var templateDocument = templateMainPart.Document;
            var templateBody = templateDocument.Body;

            Console.WriteLine("Grabbing images...");
            var templateImageParts = templateMainPart.ImageParts;

            Console.WriteLine("Creating output file...");
            using var outputWordDocument = WordprocessingDocument.Create(outputFileName, WordprocessingDocumentType.Document);
            var outputMainPart = outputWordDocument.AddMainDocumentPart();
            outputMainPart.Document = new Document(templateDocument.OuterXml);
            var outputDocument = outputMainPart.Document;
            outputDocument.Body = null;
            var outputBody = outputDocument.AppendChild(new Body(templateBody.OuterXml));
            outputBody.RemoveAllChildren<Paragraph>();

            Console.WriteLine("Adding images...");
            var templateToOutputImageParts = templateImageParts.ToDictionary(
                image => image,
                image => outputMainPart.AddPart(image));
            var templateToOutputImageIds = templateToOutputImageParts.ToDictionary(
                kvPair => templateMainPart.GetIdOfPart(kvPair.Key),
                kvPair => outputMainPart.GetIdOfPart(kvPair.Value));

            Console.WriteLine("Inserting values...");
            var alteredBodies = mapping.Select(map =>
            {
                var text = templateBody.OuterXml;
                foreach (var kvPair in map)
                {
                    text = text.Replace($"${{{kvPair.Key}}}", kvPair.Value);
                }
                foreach(var kvPair in templateToOutputImageIds)
                {
                    text = text.Replace($@"embed=""{kvPair.Key}""", $@"embed=""{kvPair.Value}""");
                }
                return new Body(text);
            });

            Console.WriteLine("Adding text...");
            foreach (var alteredBody in alteredBodies)
            {
                foreach (var child in alteredBody.ChildElements.OfType<Paragraph>())
                {
                    outputBody.AppendChild(new Paragraph(child.OuterXml));
                }
            }

            Console.WriteLine("Exporting document...");
            using var outputWriter = new StreamWriter(outputMainPart.GetStream(FileMode.Create));
            await outputWriter.FlushAsync();
        }
    }
}

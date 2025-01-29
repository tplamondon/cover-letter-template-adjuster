using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace cover_letter_template_adjuster
{
    internal class Program
    {

        internal static readonly string COMPANY_TEXT = "COMPANY_NAME";
        internal static readonly string ROLE_TEXT = "ROLE_NAME";

        /// <summary>
        /// Args[0] = file path
        /// Args[1] = company name
        /// Args[2] = position name
        /// </summary>
        /// <param name="args"></param>
        internal static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
                .WithParsed(CreateCoverLetterFromTemplate);

        }

        internal static void CreateCoverLetterFromTemplate(Options options)
        {
            //Append .docx if required
            if (!options.Path.EndsWith(".docx"))
            {
                options.Path += ".docx";
            }

            string coverLetterFile;
            try
            {
                coverLetterFile = TryCreateCoverLetterFile(options);
            }
            catch (FileNotFoundException e)
            {
                Console.Error.WriteLine(e.Message);
                return;
            }

            using WordprocessingDocument document = WordprocessingDocument.Open(coverLetterFile, true);
            try
            {
                ProcessDocument(document, options.CompanyName, options.RoleName);
                Console.WriteLine("Successfully created new cover letter from template");
            }
            catch (FileFormatException e)
            {
                Console.Error.WriteLine($"Unexpected Error. {e.Message}");
            }
            
        }

        internal static string TryCreateCoverLetterFile(Options options)
        {
            string? pathFullDirectory = new FileInfo(options.Path)?.Directory?.FullName;
            string updatedCoverLetter;
            if (Directory.Exists(pathFullDirectory) && File.Exists(options.Path))
            {
                updatedCoverLetter = $"{pathFullDirectory}/{options.CompanyName} {options.RoleName} - Cover Letter.docx";
                File.Copy(options.Path, updatedCoverLetter, true);
            }
            else
            {
                throw new FileNotFoundException($"Word document with provided path {options.Path} does not exist.\nPlease check the path of the file you wish to process");
            }
            return updatedCoverLetter;
        }
        internal static void ProcessDocument(WordprocessingDocument document, string companyName, string roleName)
        {
            var documentBody = document.MainDocumentPart?.Document.Body;
            if (documentBody == null)
            {
                throw new FileFormatException("Unexpected Error processing document body. Was null");
            }
            var paragraphs = documentBody.Elements<Paragraph>();

            foreach (var paragraph in paragraphs)
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        if (text.Text.Contains(COMPANY_TEXT))
                        {
                            text.Text = text.Text.Replace(COMPANY_TEXT, companyName);
                        }
                        if (text.Text.Contains(ROLE_TEXT))
                        {
                            text.Text = text.Text.Replace(ROLE_TEXT, roleName);
                        }
                    }
                }
            }
            if (document.CanSave)
            {
                document.Save();
            }
        }

    }
}

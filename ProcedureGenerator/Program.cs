using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ProcedureGenerator
{
    public class Procedure
    {
        public List<Person> People { get; set; }
        public List<Version> Versioning { get; set; }

        [XmlElement("Section")]
        public List<Section> Sections { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string ProjectName { get; set; }
    }

    public class Person
    {
        public string Name { get; set; }
        public string Abbreviation { get; set; }
        public string Email { get; set; }
        public string Purpose { get; set; }
    }

    public class Version
    {
        public string TextName { get; set; }
        public DateTime Date { get; set; }
        public string Editor { get; set; }
        public string Checker { get; set; }
        public string Approver { get; set; }
        public string Description { get; set; }
    }

    class ProcedureFile
    {
        public static Procedure LoadFromFile(string filename)
        {
            using (var file = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var x = new XmlSerializer(typeof(Procedure));
                return x.Deserialize(file) as Procedure;
            }
        }
    }

    public class TextEvents : PdfPageEventHelper
    {
        private DateTime _GenerationTime = DateTime.Now;
        private PdfContentByte _ContentByte;
        private BaseFont _font;
        private PdfTemplate _header;
        private PdfTemplate _footer;

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            string fontName = "Tahoma";
            //if (!FontFactory.IsRegistered(fontName))
            //{
            //    var path = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot"), "fonts", "Tahoma.ttf");
            //    FontFactory.Register(fontName);
            //}


            _font = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, true);
            _ContentByte = writer.DirectContent;
            _header = _ContentByte.CreateTemplate(100, 100);
            _footer = _ContentByte.CreateTemplate(100, 100);
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

            var phrase = new Phrase("Halliburton");
            string text = "Page " + writer.PageNumber + " of ";
            {
                _ContentByte.BeginText();
                _ContentByte.SetFontAndSize(_font, 12);
                _ContentByte.SetTextMatrix(document.PageSize.GetRight(200), document.PageSize.GetTop(45));
                _ContentByte.ShowText(text);
                _ContentByte.EndText();
                var len = _font.GetWidthPoint(text, 12);
                _ContentByte.AddTemplate(_header, document.PageSize.GetRight(180) + len, document.PageSize.GetTop(45));
            }
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);

            _header.BeginText();
            _header.SetFontAndSize(_font, 12);
            _header.SetTextMatrix(0, 0);
            _header.ShowText((writer.PageNumber).ToString());
            _header.EndText();
        }
    }

    public class Section
    {
        [XmlElement("Section")]
        public List<Section> Sections { get; set; }
        public List<string> Procedure { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string Title { get; set; }
    }

    internal static class Extensions
    {
        public static Paragraph AddCenterText(this Document document, string text, int fontSize = 12)
        {
            var p = new Paragraph(text, new Font(Font.FontFamily.HELVETICA, fontSize))
            {
                Alignment = Element.ALIGN_CENTER
            };
            document.Add(p);
            return p;
        }

        public static PdfPCell AddEcnCell(this PdfPTable table, string text, bool isCentered = true)
        {
            var cell = new PdfPCell(new Phrase(text))
            {
                HorizontalAlignment = isCentered ? 1 : 0,
                VerticalAlignment = 1,
                MinimumHeight = 20
            };
            table.AddCell(cell);
            return cell;
        }

        public static PdfPCell AddHeaderCell(this PdfPTable table, string text)
        {
            var p = new Phrase(text, new Font(Font.FontFamily.HELVETICA, 10, Font.BOLD));
            var cell = new PdfPCell(p)
            {
                HorizontalAlignment = 1,
                VerticalAlignment = 1,
                MinimumHeight = 15,
                GrayFill = 0.85f
            };
            table.AddCell(cell);
            return cell;
        }

        static Stack<int> _sectionNumber = new Stack<int>(new[] { 1 });

        public static Paragraph AddNewSection(this Document document, string text)
        {
            _sectionNumber.Push(1);
            text = string.Join(".", _sectionNumber) + " " + text;
            var p = new Paragraph(text)
            {
                Font = new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD)
            };
            document.Add(p);
            return p;
        }

        public static Paragraph AddSection(this Document document, string text)
        {
            int i = _sectionNumber.Pop();
            _sectionNumber.Push(i + 1);
            text = string.Join(".", _sectionNumber) + " " + text;
            var p = new Paragraph(text)
            {
                Font = new Font(Font.FontFamily.HELVETICA, 18, Font.BOLD)
            };
            document.Add(p);
            return p;
        }
    }

    class Program
    {

        private static void AddSections(Document document, Section section)
        {
            foreach (var subsection in section.Sections)
            {
                document.AddNewSection(subsection.Title);
                AddSections(document, subsection);
            }
        }

        private static void Main(string[] args)
        {
            var procedure = ProcedureFile.LoadFromFile(@"proc.xml");
            using (var document = new Document(PageSize.LETTER, 0.5f, 0.5f, 0.5f, 0.5f))
            using (var writer = PdfWriter.GetInstance(document, new FileStream(@"proc.pdf", FileMode.Create)))
            {
                writer.PageEvent = new TextEvents();

                document.Open();

                document.AddCenterText(procedure.Title, 24);
                document.AddCenterText(" ", 20);
                document.AddCenterText(procedure.Subtitle, 20);
                document.AddCenterText(" ", 20);
                document.AddCenterText(procedure.ProjectName, 20);

                var table = new PdfPTable(new[] { 0.3f, 0.1f, 0.8f, 0.1f, 0.1f, 0.1f });
                foreach (var version in procedure.Versioning)
                {
                    table.AddEcnCell(version.Date.ToString("yyyy-MMM-dd"));
                    table.AddEcnCell(version.TextName);
                    table.AddEcnCell(version.Description, false);
                    table.AddEcnCell(version.Editor);
                    table.AddEcnCell(version.Checker);
                    table.AddEcnCell(version.Approver);
                }

                table.AddHeaderCell("Date");
                table.AddHeaderCell("Rev");
                table.AddHeaderCell("Description");
                table.AddHeaderCell("Edit");
                table.AddHeaderCell("Chkr");
                table.AddHeaderCell("Appr");

                document.Add(table);

                document.NewPage();

                document.Add(new Paragraph("Table of Contents"));

                document.NewPage();

                foreach (var section in procedure.Sections)
                {
                    document.AddSection(section.Title ?? section.Name);
                    AddSections(document, section);
                }

                document.Close();
            }
        }
    }
}

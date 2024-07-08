using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using DocJustification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using DocumentFormat.OpenXml.ExtendedProperties;

namespace LatexToWord
{
    class SimpleWordDocument : IDisposable
    {
        private bool disposed = false;
        private WordprocessingDocument wordDocument;
        private MainDocumentPart mainPart;
        private Body body;

        private string outputFileName;

        // default parameter prop
        DocProperties paraDefaultProps;

        private Paragraph temp_paragraph;
        private Run temp_run;

        public SimpleWordDocument(string outputFileName)
        {
            this.outputFileName = outputFileName;
            wordDocument = WordprocessingDocument.Create(outputFileName, WordprocessingDocumentType.Document);

            // MainDocumentPartを追加
            mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document();
            body = new Body();
            mainPart.Document.Append(body);

            // default prop 
            paraDefaultProps = MakeDefaultProp();
        }

        /// <summary>
        /// IDIsposable implement
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// IDIsposable implement
        /// </summary>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // マネージドリソースの解放
                    if (wordDocument != null)
                    {
                        mainPart.Document.Save();
                        wordDocument.Dispose();
                        wordDocument = null;
                    }
                }
                // アンマネージドリソースの解放

                disposed = true;
            }
        }

        /// <summary>
        /// IDIsposable implement
        /// </summary>
        ~SimpleWordDocument()
        {
            Dispose(false);
        }

        public DocProperties MakeDefaultProp()
        {
            DocProperties paraProps = new DocProperties();
            DocJustification justification = new DocJustification() { Val = DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Left };
            paraProps.Append(justification);

            // 段落のスタイルを設定
            SpacingBetweenLines spacing = new SpacingBetweenLines()
            {
                LineRule = LineSpacingRuleValues.Exact,
                Line = "276"
            };
            paraProps.Append(spacing);

            return paraProps;
        }

        public DocProperties SetParagraphProp(DocProperties prop)
        {
            DocProperties oldProp = MakeDefaultProp();
            paraDefaultProps = prop;
            return oldProp;
        }

        public void AddText(Run run, string line)
        {
            run.Append(new Text(line));
        }

        public void AddParagraph(string line)
        {
            // make new paragraph
            Paragraph para = new Paragraph();

            // set style to paragraph
            // paraDefaultProps.Clone();
            ParagraphProperties prop = (ParagraphProperties)paraDefaultProps.Clone();
            para.Append(prop);

            Run run = new Run();
            run.Append(new Text(line));

            para.Append(run);

            body.Append(para);
        }


        public void AddParagraphs(string[] lines)
        {

            // make new paragraph
            Paragraph para = new Paragraph();

            // set style to paragraph
            // paraDefaultProps.Clone();
            ParagraphProperties prop = (ParagraphProperties)paraDefaultProps.Clone();
            para.Append(prop);

            Run run = new Run();
            foreach (var paragraph in lines)
            {
                AddText(run, paragraph);
                // 改行を挿入
                Break breakElement = new Break(); // 改行要素
                run.Append(breakElement);
            }

            para.Append(run);
            body.Append(para);

        }

        public void AddMath( Run run, string text)
        {
            // set style to paragraph
            // paraDefaultProps.Clone();
            // ParagraphProperties prop = (ParagraphProperties)paraDefaultProps.Clone();
            // para.Append(prop);
            OfficeMath officeMath = new OfficeMath();
            run.Append(officeMath);
            run.Append(new Text(text));
        }

        public void AddMathParagraph(string text)
        {

            // make new paragraph
            Paragraph para = new Paragraph();


            // set style to paragraph
            // paraDefaultProps.Clone();
            ParagraphProperties prop = (ParagraphProperties)paraDefaultProps.Clone();
            para.Append(prop);

            Run run = new Run();
            AddMath(run, text);
            
            para.Append(run);
            body.Append(para);
        }

        public Run BeginParagraph() { 
            temp_paragraph = new Paragraph();
            temp_run = new Run();
            return temp_run; 
        }
        public void EndParagraph()
        {
            temp_paragraph.Append(temp_run);
            body.Append(temp_paragraph);
        }


    }
}

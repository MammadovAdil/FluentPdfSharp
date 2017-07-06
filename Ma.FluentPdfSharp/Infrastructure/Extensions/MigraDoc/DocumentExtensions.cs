using Ma.FluentPdfSharp.Models.MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;

namespace Ma.FluentPdfSharp.Infrastructure.Extensions.MigraDoc
{
    public static class DocumentExtensions
    {
        /// <summary>
        /// Initialize styles for document. 
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When document is null.
        /// </exception>
        /// <param name="document">Document to initialize styles for.</param>
        /// <returns>Migradoc document.</returns>
        public static Document InitializeStyles(this Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // Get Normal style
            Style style = document.Styles.Normal;

            // Set font name            
            document.SetFont(14);

            // Add form name style
            style = document.Styles.AddStyle(
                CustomStyleNames.FormName, StyleNames.Normal);
            style.Font.Bold = true;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Right;            

            // Edit header style
            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            style.ParagraphFormat.SpaceBefore = Unit.FromCentimeter(0.5);            
            style.Font.Bold = true;
            style.Font.Size = 16;

            // Edit footer style
            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            style.ParagraphFormat.Borders.Top.Style = BorderStyle.DashLargeGap;
            style.Font.Size = 12;
            style.Font.Italic = true;
            style.Font.Color = Color.FromRgbColor(150, Colors.Black); 

            // Edit first heading style
            style = document.Styles[StyleNames.Heading1];
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.SpaceBefore = Unit.FromCentimeter(1);
            style.ParagraphFormat.SpaceAfter = Unit.FromCentimeter(0);                  
            style.Font.Bold = true;
            style.Font.Size = 16;

            // Edit third heading style
            style = document.Styles[StyleNames.Heading3];
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.ParagraphFormat.SpaceBefore = Unit.FromCentimeter(0);
            style.ParagraphFormat.SpaceAfter = Unit.FromCentimeter(0);
            style.Font.Bold = true;
            style.Font.Size = 12;

            // Add field paragraph style
            style = document.Styles.AddStyle(
                CustomStyleNames.FieldParagraph, StyleNames.Normal);
            style.ParagraphFormat.SpaceBefore = Unit.FromCentimeter(-0.5);
            style.ParagraphFormat.SpaceAfter = 0;            

            // Add label style
            style = document.Styles.AddStyle(
                CustomStyleNames.FieldLabel, StyleNames.Normal);
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
            style.Font.Color = Color.FromRgbColor(180, Colors.Black);    

            // Add field style
            style = document.Styles.AddStyle(
                CustomStyleNames.FieldValue, StyleNames.Normal);

            // Add bold style   
            style = document.Styles.AddStyle(
                CustomStyleNames.Bold, StyleNames.Normal);
            style.Font.Bold = true;

            // Add bottom border style
            style = document.Styles.AddStyle(
                CustomStyleNames.BottomBorderTable, StyleNames.Normal);            
            style.ParagraphFormat.Borders.Bottom.Style = BorderStyle.Single;

            return document;           
        }

        /// <summary>
        /// Set font for document.
        /// </summary>
        /// <param name="document">Document to </param>
        /// <param name="fontSize">Size of font.</param>
        /// <returns>MigraDoc document.</returns>
        public static Document SetFont(this Document document, Unit fontSize)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // Get Normal style
            Style style = document.Styles.Normal;

            // Set font name            
            style.Font.Name = "Segoe UI";
            style.Font.Size = fontSize;

            return document;
        }

        /// <summary>
        /// Add pages of MigraDoc document to PdfSharp document.
        /// </summary>
        /// <param name="document">MigraDoc document.</param>
        /// <param name="pdfSharpDocument">PdfSharp document.</param>
        public static void MergeInto(
            this Document document,
            PdfDocument pdfSharpDocument)
        {
            DocumentRenderer docRenderer = new DocumentRenderer(document);
            docRenderer.PrepareDocument();

            for (int i = 0; i < docRenderer.FormattedDocument.PageCount; i++)
            {
                // Add new page to pdf sharp document.
                PdfPage page = pdfSharpDocument.AddPage();
                page.Size = PdfSharp.PageSize.A4;

                XGraphics gfx = XGraphics.FromPdfPage(page);
                gfx.MUH = PdfFontEncoding.Unicode;
                gfx.MFEH = PdfFontEmbedding.Default;

                // Render MigraDoc page to PdfSharp page.
                // Remember that page number start with 1.
                docRenderer.RenderPage(gfx, i + 1);
            }            
        }
    }
}

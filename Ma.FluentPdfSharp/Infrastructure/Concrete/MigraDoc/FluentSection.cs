using Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc;
using Ma.FluentPdfSharp.Models.MigraDoc;
using MigraDoc.DocumentObjectModel;
using System;
using Ma.FluentPdfSharp.Infrastructure.Extensions.MigraDoc;

namespace Ma.FluentPdfSharp.Infrastructure.Concrete.MigraDoc
{
    public class FluentSection : IFluentSection
    {
        /// <summary>
        /// Underlying section.
        /// </summary>
        public Section Section { get; private set; }

        /// <summary>
        /// Initialize fluent section.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When sectionParam is null.
        /// </exception>
        /// <param name="sectionParam">Underlying section.</param>
        public FluentSection(Section sectionParam)
        {
            if (sectionParam == null)
                throw new ArgumentNullException(nameof(sectionParam));

            Section = sectionParam;
        }

        /// <summary>
        /// Add paragraph to section.
        /// </summary>
        /// <param name="text">Text of paragraph.</param>
        /// <returns>Insrance of IFluentSection.</returns>
        public IFluentSection AddParagraph(string text)
        {
            Section.AddParagraph(text);

            return this;
        }

        /// <summary>
        /// Add paragraph to section.
        /// </summary>
        /// <param name="text">Text of paragraph.</param>
        /// <param name="style">Style of paragraph.</param>
        /// <returns>Insrance of IFluentSection.</returns>
        public IFluentSection AddParagraph(string text, string style)
        {
            Section.AddParagraph(text, style);

            return this;
        }

        /// <summary>
        /// Add field paragraph with formatted label and text.
        /// </summary>
        /// <param name="label">Label for the field.</param>
        /// <param name="text">Text of the field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentSection AddFieldParagraph(string label, string text)
        {
            // Add paragraph with FieldParagraph style
            Paragraph fieldParagraph = Section.AddParagraph();
            fieldParagraph.Style = CustomStyleNames.FieldParagraph;

            // Add label and text
            fieldParagraph.AsFluent()
                .AddLabelText(label)
                .AddFieldText(text);

            return this;
        }
    }
}

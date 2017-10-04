using Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc;
using Ma.FluentPdfSharp.Models.MigraDoc;
using MigraDoc.DocumentObjectModel;
using System;

namespace Ma.FluentPdfSharp.Infrastructure.Concrete.MigraDoc
{
    public class FluentParagraph
        : IFluentParagraph
    {
        /// <summary>
        /// Underlying paragraph.
        /// </summary>
        public Paragraph Paragraph { get; private set; }

        /// <summary>
        /// Initialize fluent paragraph.
        /// </summary>
        /// <param name="paragraphParam">Underlying paragraph object.</param>
        public FluentParagraph(Paragraph paragraphParam)
        {
            if (paragraphParam == null)
                throw new ArgumentNullException(nameof(paragraphParam));

            Paragraph = paragraphParam;
        }

        /// <summary>
        /// Add text to paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddText(string text)
        {
            Paragraph.AddText(text);

            return this;
        }

        /// <summary>
        /// Add formatted text to paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <param name="styleName">Name of style to apply to text.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddFormattedText(string text, string styleName)
        {
            Paragraph.AddFormattedText(text, styleName);

            return this;
        }

        /// <summary>
        /// Add tab indent.
        /// </summary>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddTab()
        {
            Paragraph.AddTab();

            return this;
        }

        /// <summary>
        /// Add line break.
        /// </summary>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddLineBreak()
        {
            Paragraph.AddLineBreak();

            return this;
        }

        /// <summary>
        /// Add text as label to paragraph.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When text is null or empty.
        /// </exception>
        /// <param name="label">Text to add as label.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddLabelText(string label)
        {
            if (string.IsNullOrEmpty(label))
                throw new ArgumentNullException(nameof(label));

            // Add ':' to the end of label
            label = string.Format("{0}:", label.TrimEnd(':'));

            AddFormattedText(label, CustomStyleNames.FieldLabel);

            return this;
        }

        /// <summary>
        /// Add text as field to parahraph.
        /// </summary>
        /// <param name="text">Text to add as field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddFieldText(string text)
        {
            // If no text to add as a field, return
            if (string.IsNullOrEmpty(text))
                return this;

            AddFormattedText(text, CustomStyleNames.FieldValue);

            return this;
        }

        /// <summary>
        /// Add field with label.
        /// </summary>
        /// <param name="label">Label for the field.</param>
        /// <param name="text">Value of the field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph AddField(string label, string text)
        {
            AddLineBreak()
                .AddLabelText(label)
                .AddText(" ")
                .AddFieldText(text);

            return this;
        }

        /// <summary>
        /// Set style of paragraph.
        /// </summary>
        /// <param name="style">Style to set for paragraph.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public IFluentParagraph WithStyle(string style)
        {
            Paragraph.Style = style;

            return this;
        }
    }
}

using MigraDoc.DocumentObjectModel;

namespace Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc
{
    public interface IFluentParagraph
    {
        /// <summary>
        /// Underlying paragraph.
        /// </summary>
        Paragraph Paragraph { get; }

        /// <summary>
        /// Add text to paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddText(string text);

        /// <summary>
        /// Add formatted text to paragraph.
        /// </summary>
        /// <param name="text">Text to add.</param>
        /// <param name="styleName">Name of style to apply to text.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddFormattedText(string text, string styleName);

        /// <summary>
        /// Add tab indent.
        /// </summary>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddTab();

        /// <summary>
        /// Add line break.
        /// </summary>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddLineBreak();

        /// <summary>
        /// Add text as label to paragraph.
        /// </summary>
        /// <param name="label">Text to add as label.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddLabelText(string label);

        /// <summary>
        /// Add text as field to parahraph.
        /// </summary>
        /// <param name="text">Text to add as field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddFieldText(string text);

        /// <summary>
        /// Add field with label.
        /// </summary>
        /// <param name="label">Label for the field.</param>
        /// <param name="text">Value of the field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph AddField(string label, string text);

        /// <summary>
        /// Set style of paragraph.
        /// </summary>
        /// <param name="style">Style to set for paragraph.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentParagraph WithStyle(string style);
    }
}

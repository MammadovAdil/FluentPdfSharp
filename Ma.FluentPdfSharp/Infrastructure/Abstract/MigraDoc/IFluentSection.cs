using MigraDoc.DocumentObjectModel;

namespace Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc
{
    public interface IFluentSection
    {
        /// <summary>
        /// Underlying section.
        /// </summary>
        Section Section { get; }

        /// <summary>
        /// Add paragraph to section.
        /// </summary>
        /// <param name="text">Text of paragraph.</param>
        /// <returns>Insrance of IFluentSection.</returns>
        IFluentSection AddParagraph(string text);

        /// <summary>
        /// Add paragraph to section.
        /// </summary>
        /// <param name="text">Text of paragraph.</param>
        /// <param name="style">Style of paragraph.</param>
        /// <returns>Insrance of IFluentSection.</returns>
        IFluentSection AddParagraph(string text, string style);

        /// <summary>
        /// Add field paragraph with formatted label and text.
        /// </summary>
        /// <param name="label">Label for the field.</param>
        /// <param name="text">Text of the field.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        IFluentSection AddFieldParagraph(string label, string text);
    }
}

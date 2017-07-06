using Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc;
using Ma.FluentPdfSharp.Infrastructure.Concrete.MigraDoc;
using MigraDoc.DocumentObjectModel;
using System;

namespace Ma.FluentPdfSharp.Infrastructure.Extensions.MigraDoc
{
    public static class ParagraphExtensions
    {
        /// <summary>
        /// Get fluent wrapper for paragraph.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When paragraph is null.
        /// </exception>
        /// <param name="paragraph">Paragraph to get fluent wrapper for.</param>
        /// <returns>Instance of IFluentParagraph.</returns>
        public static IFluentParagraph AsFluent(
            this Paragraph paragraph)
        {
            if (paragraph == null)
                throw new ArgumentNullException(nameof(paragraph));

            IFluentParagraph fluentParagraph = new FluentParagraph(paragraph);
            return fluentParagraph;
        }
    }
}

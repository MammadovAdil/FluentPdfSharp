using Ma.FluentPdfSharp.Infrastructure.Abstract.MigraDoc;
using Ma.FluentPdfSharp.Infrastructure.Concrete.MigraDoc;
using MigraDoc.DocumentObjectModel;
using System;

namespace Ma.FluentPdfSharp.Infrastructure.Extensions.MigraDoc
{
    public static class SectionExtensions
    {
        /// <summary>
        /// Get fluent wrapper for section.
        /// </summary>
        /// <exception cref="ArgumentNullException">
        /// When section is null.
        /// </exception>
        /// <param name="section">Underlying section.</param>
        /// <returns>Instance of IFluentSection.</returns>
        public static IFluentSection AsFluent(this Section section)
        {
            if (section == null)
                throw new ArgumentNullException(nameof(section));

            // Intiialize fluent section wrapper
            IFluentSection fluentSection = new FluentSection(section);
            return fluentSection;
        }
    }
}

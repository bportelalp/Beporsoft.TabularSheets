using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Beporsoft.TabularSheets.Exceptions
{
    /// <summary>
    /// Represent an exception raised after validating the OpenXml Structure
    /// </summary>
    public class OpenXmlException : Exception
    {
        internal OpenXmlException(IEnumerable<ValidationErrorInfo> errors) : base("Errors occurred during validation of the OpenXml document")
        {
            foreach (var error in errors)
            {
                Errors.Add(new ValidationError()
                {
                    Message = error.Description
                });
            }
            OriginalErrors = errors;

        }

        /// <summary>
        /// The collection of errors during validation
        /// </summary>
        public ICollection<ValidationError> Errors { get; } = new List<ValidationError>();

        /// <summary>
        /// The original list of erros
        /// </summary>
        internal IEnumerable<ValidationErrorInfo> OriginalErrors { get; }

        /// <summary>
        /// Information of every error during validation
        /// </summary>
        public record ValidationError
        {
            /// <summary>
            /// Message of the error
            /// </summary>
            public string Message { get; set; } = null!;

        }

    }
}

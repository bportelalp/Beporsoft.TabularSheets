using Beporsoft.TabularSheets.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets.Builders
{
    internal class SpreadsheetValidator
    {
        public static IEnumerable<ValidationErrorInfo> Validate(SpreadsheetDocument document)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            IEnumerable<ValidationErrorInfo> errors = validator.Validate(document);
            return errors;
        }
    }
}

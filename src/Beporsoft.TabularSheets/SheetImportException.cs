using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    public class SheetImportException : Exception
    {
        public SheetImportException(string message) : base(message)
        {

        }

        public SheetImportException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}

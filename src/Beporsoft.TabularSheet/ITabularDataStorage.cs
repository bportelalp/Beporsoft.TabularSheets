using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheet
{
    /// <summary>
    /// Interface which grants methods to create tables
    /// </summary>
    internal interface ITabularDataStorage
    {
        /// <summary>
        /// Create the tabular data on a file support.
        /// </summary>
        /// <param name="path">The path to file</param>
        public void Create(string path);

        /// <summary>
        /// Create the tabular data as memory stream.
        /// </summary>
        /// <returns></returns>
        public MemoryStream Create();
    }
}

using Beporsoft.TabularSheets.Builders;
using Beporsoft.TabularSheets.Builders.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Beporsoft.TabularSheets
{
    internal sealed class TabularBook : IList<ISheet>
    {
        private List<ISheet> _sheets = new();

        public TabularBook()
        {
            
        }
        public IList<ISheet> Sheets { get => _sheets; set => _sheets = value.ToList(); }

        /// <summary>
        /// Gets the number of rows.
        /// </summary>
        public int Count => _sheets.Count;

        /// <inheritdoc cref="ICollection{T}.IsReadOnly"/>
        public bool IsReadOnly => false;

        #region IList<ISheet>
        public ISheet this[int index] { get => _sheets[index]; set => _sheets[index] = value; }


        /// <inheritdoc cref="ICollection{T}.Add(T)"/>
        public void Add(ISheet sheet) => _sheets.Add(sheet);

        /// <inheritdoc cref="ICollection{T}.Clear"/>
        public void Clear() => _sheets.Clear();

        /// <inheritdoc cref="ICollection{T}.Contains(T)"/>
        public bool Contains(ISheet item) => _sheets.Contains(item);

        /// <inheritdoc cref="ICollection{T}.CopyTo(T[], int)"/>
        public void CopyTo(ISheet[] array, int arrayIndex) => _sheets.CopyTo(array, arrayIndex);

        /// <inheritdoc cref="IEnumerable{T}.GetEnumerator"/>
        public IEnumerator<ISheet> GetEnumerator() => _sheets.GetEnumerator();

        /// <inheritdoc cref="IList{T}.IndexOf(T)"/>
        public int IndexOf(ISheet item) => _sheets.IndexOf(item);

        /// <inheritdoc cref="IList{T}.Insert(int, T)"/>
        public void Insert(int index, ISheet item) => _sheets.Insert(index, item);

        /// <inheritdoc cref="ICollection{T}.Remove(T)"/>
        public bool Remove(ISheet item) => _sheets.Remove(item);

        public void RemoveAt(int index) => _sheets.RemoveAt(index);

        /// <inheritdoc cref="IEnumerable.GetEnumerator"/>
        IEnumerator IEnumerable.GetEnumerator() => _sheets.GetEnumerator();


        /// <inheritdoc cref="List{T}.AddRange(IEnumerable{T})"/>
        public void AddRange(IEnumerable<ISheet> items) => _sheets.AddRange(items);
        #endregion

        #region Create
        public void Create(string path)
        {
        }

        public MemoryStream Create()
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}

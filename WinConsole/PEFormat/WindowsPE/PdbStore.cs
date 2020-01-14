using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsPE
{
    [ComVisible(true)]
    [Guid("34B66A2A-F694-49A7-A48F-3C7BDFD363C5")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPdbStore
    {
        void Add(SYMBOL_INFO si);
    }

    [ComVisible(true)]
    [Guid("7B6BC40A-AF93-4748-8F50-4F81A1FF4BBB")]
    [ClassInterface(ClassInterfaceType.None)]
    public class PdbStore : IPdbStore
    {
        Dictionary<string, SYMBOL_INFO> _cache = new Dictionary<string, SYMBOL_INFO>();

        public PdbStore() { }

        public void Add(SYMBOL_INFO si)
        {
            if (_cache.ContainsKey(si.Name) == false) // duplicates of "<anonymous-tag>"
            {
                _cache.Add(si.Name, si);
            }
        }

        public SYMBOL_INFO this[string key]
        {
            get
            {
                if (_cache.ContainsKey(key) == false)
                {
                    return default(SYMBOL_INFO);
                }

                return _cache[key];
            }
        }

        public IEnumerable<SYMBOL_INFO> Enumerate()
        {
            foreach (var item in _cache.Values)
            {
                yield return item;
            }
        }
    }
}

#if NETSTANDARD1_3
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace Excel
{
    public class PropertyCollection : IEnumerable<KeyValuePair<string, string>>
    {
        private readonly Dictionary<string, string> dict = new Dictionary<string, string>();

        public string this[string key] => dict[key];

        public void Add(string key, string value)
        {
            dict.Add(key, value);
        }

        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            return dict.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return dict.GetEnumerator();
        }
    }
}
#endif
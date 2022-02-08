namespace PflanzenPaesse.Repositories.ConstantsRepository
{
    using System.Collections;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Linq;

    public static class ConstantsRepository
    {
        public static IDictionary<string, string> Import()
        {
            var section = (Hashtable)ConfigurationManager.GetSection("Constants");
            return section.Cast<DictionaryEntry>().ToDictionary(x => x.Key.ToString(), x => x.Value.ToString());
        }
    }
}

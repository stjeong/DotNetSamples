using System;
using System.DirectoryServices;

namespace LdapObjectSearch
{
    class Program
    {
        static void Main(string[] args)
        {        // string path = "LDAP://CN=testuser,CN=Users,DC=testad,DC=com";
                 // string path = "LDAP://CN=testuser,OU=TestGrp,DC=testad,DC=com";
            string path = "LDAP://CN=Guests,CN=Builtin,DC=testad,DC=com";

            DirectoryEntry searchRoot = new DirectoryEntry(path);
            DirectorySearcher search = new DirectorySearcher(searchRoot);

            try
            {
                foreach (SearchResult resEnt in search.FindAll())
                {
                    WriteAttr(resEnt, "cn");
                    WriteAttr(resEnt, "distinguishedName");
                }
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private static void WriteAttr(SearchResult result, string attrName)
        {
            if (result.Properties.Contains(attrName) == true)
            {
                Console.WriteLine(attrName + " == " + (String)result.Properties[attrName][0]);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Security.Principal;
using System.DirectoryServices;

namespace RegKor.Classess
{
    public class MyAplicationIdentity
    {
        /// <summary>
        /// Возвращает реальное имя пользователя в домене
        /// </summary>
        /// <returns></returns>
        public static string GetUses()
        {
            WindowsIdentity id = WindowsIdentity.GetCurrent();
            string domainName = id.Name.Split('\\')[0];
            string userName = id.Name.Split('\\')[1];
            string Sid = id.User.Value.ToString();
            string ldapPath = "LDAP://<SID=" + Sid + ">";

            DirectoryEntry g = new DirectoryEntry(ldapPath);

            g.UsePropertyCache = true;
            DirectorySearcher ds = new DirectorySearcher(g);
            string userFilter = "(&(objectClass=user)(objectCategory=Person)(sAMAccountName={0}))";

            ds.SearchScope = SearchScope.Subtree;
            ds.PropertiesToLoad.Add("cn");
            ds.PageSize = 1;
            ds.ServerPageTimeLimit = TimeSpan.FromSeconds(2);
            ds.Filter = string.Format(userFilter, userName);

            SearchResult sr = ds.FindOne();
            string cn = (string)sr.Properties["cn"][0];

            return cn;
        }

    }
}

using EH.System.Models.Dtos;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;


namespace EH.System.Commons
{
    //public interface IADHelper : IScoped
    //{
    //    void Send(string formEmail, string toEmail, string subject, string body);

    //}
    public class ADHelper : IScoped
    {
        #region //   

        #region Server and Admin Info
        // LDAP地址 例如：LDAP://my.com.cn
        private readonly string LDAP_HOST; //= "LDAP://RootDSE";
        // 具有LDAP管理权限的特殊帐号
        private readonly string AdminUSER;// = "paceyl";
        // 具有LDAP管理权限的特殊帐号的密码
        private readonly string AdminPASSWORD;//

        private readonly IConfiguration configuration;
        private readonly LogHelper logHelper;

        public ADHelper(IConfiguration configuration, LogHelper logHelper)
        {
            string key = "5ehealth";
            this.configuration = configuration;
            this.logHelper = logHelper;

            var host = configuration.GetSection("ADSetting:Host").Value.ToString();

            var adminUser = (configuration.GetSection("ADSetting:User").Value);
            //var s = EncryptHelper.DESEncrypt("ebdt~(3724pom",key,key);
            var password = EncryptHelper.DESDecrypt(configuration.GetSection("ADSetting:Password").Value, key, key);
            LDAP_HOST = host;
            AdminUSER = adminUser;
            AdminPASSWORD = password;
        }

        #endregion

        //#region MyRegion
        ///// <summary>
        ///// 获取组织单位
        ///// </summary>
        ///// <param name="parent"></param>
        ///// <param name="ouname"></param>
        ///// <returns></returns>
        //public static DirectoryEntry GetOU(DirectoryEntry parent, string ouname)
        //{
        //    using (DirectorySearcher mySearcher = new DirectorySearcher(parent, "(objectclass=organizationalUnit)"))
        //    {
        //        using (DirectorySearcher deSearch = new DirectorySearcher())
        //        {
        //            deSearch.SearchRoot = parent;
        //            deSearch.Filter = string.Format("(&(objectClass=organizationalUnit) (OU={0}))", ouname);
        //            SearchResult results = deSearch.FindOne();
        //            if (results != null)
        //            {
        //                return results.GetDirectoryEntry();
        //            }
        //            else
        //            {
        //                return null;
        //            }
        //        }
        //    }
        //}

        ///// <summary>
        ///// 建组织单位
        ///// </summary>
        ///// <param name="parent"></param>
        ///// <param name="ouname"></param>
        //public static void AddOU(DirectoryEntry parent, string ouname)
        //{
        //    DirectoryEntries ous = parent.Children;
        //    DirectoryEntry ou = ous.Add("OU=" + ouname, "organizationalUnit");
        //    ou.CommitChanges();
        //    ou.Close();
        //}

        ///// <summary>
        ///// 建立连接
        ///// </summary>
        ///// <param name="oupath"></param>
        ///// <returns></returns>
        //public static PrincipalContext createConnection(List<string> oupath = null)
        //{
        //    string path = "";
        //    foreach (string str in _domainArr)
        //    {
        //        path += string.Format(",DC={0}", str);
        //    }
        //    if (oupath != null)
        //    {
        //        string tmp = "";
        //        for (int i = oupath.Count - 1; i >= 0; i--)
        //        {
        //            tmp += string.Format(",OU={0}", oupath[i]);
        //        }
        //        tmp = tmp.Substring(1);
        //        path = tmp + path;
        //    }
        //    else
        //    {
        //        path = path.Substring(1);
        //    }

        //    var context = new PrincipalContext(ContextType.Domain, _domain, path, ContextOptions.Negotiate, AdminUSER, AdminPASSWORD);
        //    return context;
        //}

        ///// <summary>
        ///// 添加用户
        ///// </summary>
        ///// <param name="context"></param>
        ///// <param name="barcode"></param>
        ///// <param name="userName"></param>
        ///// <param name="passWord"></param>
        //public static void AddUser(PrincipalContext context, string barcode, string userName, string passWord)
        //{
        //    using (UserPrincipal u = new UserPrincipal(context, barcode, passWord, true))
        //    {
        //        u.Name = barcode;
        //        u.DisplayName = userName;
        //        u.UserCannotChangePassword = true;
        //        u.PasswordNotRequired = true;
        //        u.PasswordNeverExpires = true;
        //        u.UserPrincipalName = barcode + "@" + _domain;
        //        u.Save();
        //    }
        //}

        ///// <summary>
        ///// 删除用户
        ///// </summary>
        ///// <param name="userName"></param>
        //public static void DelUser(string userName)
        //{
        //    using (var context = createConnection())
        //    {
        //        UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);
        //        if (user != null)
        //        {
        //            user.Delete();
        //        }
        //    }
        //}

        ///// <summary>
        ///// 修改密码
        ///// </summary>
        ///// <param name="userName"></param>
        ///// <param name="passWord"></param>
        //public static void EditPass(string userName, string passWord)
        //{
        //    using (var context = createConnection())
        //    {
        //        UserPrincipal user = UserPrincipal.FindByIdentity(context, userName);
        //        if (user != null)
        //        {
        //            user.SetPassword(passWord);
        //            user.Save();
        //        }
        //    }
        //}

        ///// <summary>
        ///// 登陆验证
        ///// </summary>
        ///// <param name="name"></param>
        ///// <param name="password"></param>
        ///// <returns></returns>
        //public static bool login(string name, string password)
        //{
        //    DirectoryEntry root = null;
        //    try
        //    {
        //        string ADPath = rootPath();
        //        root = new DirectoryEntry(ADPath, name, password, AuthenticationTypes.Secure);
        //        string strName = root.Name;
        //        root.Close();
        //        root = null;
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine(ex.Message);
        //        return false;
        //    }
        //}
        //#endregion


        /// <summary>
        /// 获取DefaultNamingContext属性的值
        /// </summary>
        /// <param name="host"></param>
        /// <returns></returns>
        public string GetDefaultNamingContext()
        {
            using (DirectoryEntry objRootDSE = new(LDAP_HOST, AdminUSER, AdminPASSWORD))
            {
                // 获取DefaultNamingContext属性的值
                string strDomain = objRootDSE.Properties["defaultNamingContext"].Value.ToString();
                string strLDP = "LDAP://" + strDomain;
                return strLDP;
            }
        }
        /// <summary>
        /// 验证登录
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public bool CheckLogin(string userName, string password)
        {
            var domain = GetDefaultNamingContext();
            try
            {
                using (DirectoryEntry userEntry = new DirectoryEntry(domain, userName, password))
                {
                    var result = userEntry.NativeObject;
                    if (result != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception ex)
            {

                return false;
            }

        }

        /**
         * 向某个组添加人员
         * groupName 组名称
         * userName 人员域帐号
         **/
        public void AddGroupMember(string groupName, string userName)
        {
            DirectoryEntry group = GetGroupByName(groupName);
            group.Username = AdminUSER;
            group.Password = AdminPASSWORD;
            group.Properties["member"].Add(GetUserDNByName(userName));
            group.CommitChanges();
        }

        /**
         * 从某个组移出指定的人员
         * groupName 组名称
         * userName 人员域帐号
         **/
        public void RemoveGroupMember(string groupName, string userName)
        {
            DirectoryEntry group = GetGroupByName(groupName);
            group.Username = AdminUSER;
            group.Password = AdminPASSWORD;
            group.Properties["member"].Remove(GetUserDNByName(userName));
            group.CommitChanges();
        }

        /**
         * 获取指定人员的域信息
         * name 人员域帐号 
         **/
        public object GetUserDNByName(string name)
        {
            //bool? res = false;
            //string key = "5ehealth";
            //var domain = configuration.GetSection("ADSetting:Domain").Value.ToString();
            //var admin = configuration.GetSection("ADSetting:User").Value.ToString();
            //var adPassword = EncryptHelper.DESDecrypt(configuration.GetSection("ADSetting:Password").Value, key, key);
            //using (var content = new PrincipalContext(ContextType.Domain, domain, admin, adPassword))
            //{
            //    UserPrincipal user = UserPrincipal.FindByIdentity(content, IdentityType.SamAccountName, name);
            //    if (user != null)
            //    {
            //         res = user.Enabled; 
            //    }

            //};

            //return res;
            var domain = GetDefaultNamingContext();

            DirectoryEntry entry = new DirectoryEntry(domain);
            DirectorySearcher search = new DirectorySearcher(entry);

            search.Filter = "(SAMAccountName=" + name + ")";

            SearchResult user = search.FindOne();

            if (user == null)
            {
                return null;
            }
            Dictionary<string, string> dic = new Dictionary<string, string>();
            foreach (var item in user.Properties.PropertyNames)
            {
                var key = item.ToString();
                //if (key == "distinguishedname")
                //    continue;
                //if (key == "mobile")
                //    continue;
                var value = user.Properties[key][0].ToString();
                dic[key] = value;
            }

            return JsonConvert.SerializeObject(dic);
        }

        /**
        * 获取指定人员的域信息
        * name 人员域帐号 
        **/
        public SearchResult GetUserDNByUserName(string name)
        {
            var domain = GetDefaultNamingContext();

            DirectoryEntry entry = new DirectoryEntry(domain);
            DirectorySearcher search = new DirectorySearcher(entry);

            search.Filter = "(SAMAccountName=" + name + ")";

            SearchResult user = search.FindOne();
 
            return user;
        }

        /**
         * 获取指定域组的信息
         * name 组名称 
         **/
        public DirectoryEntry GetGroupByName(string name)
        {
            var domain = GetDefaultNamingContext();

            DirectoryEntry entry = new DirectoryEntry(domain);
            DirectorySearcher search = new DirectorySearcher(entry);

            search.Filter = "(&(cn=" + name + ")(objectClass=group))";
            search.PropertiesToLoad.Add("objectClass");
            SearchResult result = search.FindOne();
            DirectoryEntry group;
            if (result != null)
            {
                group = result.GetDirectoryEntry();
            }
            else
            {
                throw new Exception("请确认AD组列表是否正确");
            }
            return group;
        }

        /// <summary>
        /// 重置密码
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        public string ResetPassword(string userName)
        {
            var domain = GetDefaultNamingContext();
            using (DirectoryEntry entry = new DirectoryEntry(domain))
            {
                try
                {
                    using (DirectorySearcher search = new DirectorySearcher(entry))
                    {
                        search.Filter = "(SAMAccountName=" + userName + ")";
                        SearchResult user = search.FindOne();

                        if (user != null)
                        {
                            DirectoryEntry userEntry = user.GetDirectoryEntry();
                            string newPassword = StringHelper.GetRandomLopLetter(4) + StringHelper.GetRandomSymbol(2) + StringHelper.GetRandomNumber(4) + StringHelper.GetRandomLopLetter(3);
                            // string newPassword = "Lpy@755815231";
                            userEntry.Invoke("SetPassword", new object[] { newPassword });
                            userEntry.CommitChanges();
                            userEntry.Close();
                            return newPassword;
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 获取厦门用户（目前确认Ad域为）
        /// LDAP://OU=Users,OU=XM,OU=EHI,DC=ehi,DC=ehealth,DC=com
        /// LDAP://OU=Citrix_Users,OU=Users,OU=XM,OU=EHI,DC=ehi,DC=ehealth,DC=com
        /// </summary>
        /// <returns></returns>
        public List<SearchResultCollection> GetUserSearchByLocation(string location)
        {
            List<SearchResultCollection> allResults = new List<SearchResultCollection>();
            SearchResultCollection userResults = null;
            var domain = GetDefaultNamingContext();
            using (DirectoryEntry entry = new DirectoryEntry(domain))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectclass=organizationalUnit)";
                    //logHelper.LogInfo(searcher.ToJson());
                    //searcher.PropertiesToLoad.Add("samaccountname");
                    //searcher.PropertiesToLoad.Add("cn");
                    //searcher.PropertiesToLoad.Add("department");
                    //searcher.PropertiesToLoad.Add("mail");
                    //searcher.PropertiesToLoad.Add("title");
                    //searcher.PropertiesToLoad.Add("manager");
                    //searcher.PropertiesToLoad.Add("userAccountControl");
                    //searcher.PropertiesToLoad.Add("whencreated");
                    //searcher.PropertiesToLoad.Add("telephonenumber"); 
                    //searcher.PropertiesToLoad.Add("pwdLastSet");

                    SearchResultCollection results = searcher.FindAll();

                    var userGroup = configuration.GetSection("ADSetting:XMUserGroup").GetChildren().Select(i => i.Value).ToArray();
                    if (location == "CN")
                    {
                        userGroup = configuration.GetSection("ADSetting:XMUserGroup").GetChildren().Select(i => i.Value).ToArray();
                    }
                    else if (location == "US")
                    {
                        userGroup = configuration.GetSection("ADSetting:USUserGroup").GetChildren().Select(i => i.Value).ToArray();
                    }
                    else
                    {
                        userGroup = configuration.GetSection("ADSetting:ALLUserGroup").GetChildren().Select(i => i.Value).ToArray();
                    }

                    foreach (SearchResult result in results)
                    {
                        //if (result.Path.Contains("OU=XM") && result.Path.Contains("OU=Users"))
                        //    searches.Add(result);

                        if (userGroup.Contains(result.Path))
                        {
                            using (DirectoryEntry userEntry = new DirectoryEntry(result.Path))
                            {
                                using (DirectorySearcher userSearch = new DirectorySearcher(userEntry))
                                {
                                    //userSearch.Filter = "(&(objectCategory=person) (objectClass=user))";
                                    userSearch.Filter = "(objectClass=user)";

                                    userSearch.PropertiesToLoad.Add("pwdlastset");
                                    userSearch.PropertiesToLoad.Add("samaccountname");
                                    userSearch.PropertiesToLoad.Add("cn");
                                    userSearch.PropertiesToLoad.Add("department");
                                    userSearch.PropertiesToLoad.Add("mail");
                                    userSearch.PropertiesToLoad.Add("title");
                                    userSearch.PropertiesToLoad.Add("manager");
                                    userSearch.PropertiesToLoad.Add("userAccountControl");
                                    userSearch.PropertiesToLoad.Add("telephonenumber");

                                    userResults = userSearch.FindAll();
                                    allResults.Add(userResults);
                                }
                            } ;
                        }
                    }
                    return allResults;
                }
            }
        }


        public SearchResultCollection GetAllUser()
        {
            SearchResultCollection userResults = null;
            var domain = GetDefaultNamingContext();
            using (DirectoryEntry entry = new DirectoryEntry(domain))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = "(objectclass=organizationalUnit)";
                    searcher.PropertiesToLoad.Add("samaccountname");
                    searcher.PropertiesToLoad.Add("cn");
                    searcher.PropertiesToLoad.Add("department");
                    searcher.PropertiesToLoad.Add("mail");
                    searcher.PropertiesToLoad.Add("title");
                    searcher.PropertiesToLoad.Add("manager");
                    searcher.PropertiesToLoad.Add("userAccountControl");
                    searcher.PropertiesToLoad.Add("whencreated");
                    searcher.PropertiesToLoad.Add("telephonenumber");
                    searcher.PropertiesToLoad.Add("whenchanged");
                    SearchResultCollection results = searcher.FindAll();

                    List<string> strings = new List<string>();
                    foreach (SearchResult result in results)
                    {
                        if (result.Path.Contains("XM") && result.Path.Contains("Users"))
                        {
                            strings.Add(result.Path);
                        }
                        //if (result.Path.Contains("OU=XM") && result.Path.Contains("OU=Users"))
                        //    searches.Add(result);
                        if (result.Path == "LDAP://OU=Users,OU=XM,OU=EHI,DC=ehi,DC=ehealth,DC=com")
                        {
                            using (DirectoryEntry userEntry = new DirectoryEntry(result.Path))
                            {
                                using (DirectorySearcher userSearch = new DirectorySearcher(userEntry))
                                {
                                    userSearch.Filter = "(objectclass=user)";
                                    userResults = userSearch.FindAll();
                                    continue;
                                    return userResults;
                                }
                            }

                        }
                    }
                    if (strings.Count > 0)
                    {

                    }
                    return userResults;
                }
            }
        }


        public long GetMaxAge()
        {
            var domain = GetDefaultNamingContext();
            long maxPwdAge = 0;
            using (DirectoryEntry entry = new DirectoryEntry(domain))
            {
                using (DirectorySearcher directorySearcher = new DirectorySearcher(entry))
                {
                    directorySearcher.SearchScope = SearchScope.Base;
                    directorySearcher.Filter = @"(objectClass=*)";
                    directorySearcher.PropertiesToLoad.Add("maxPwdAge");
                    SearchResult ouResult = directorySearcher.FindOne();
                    if (ouResult.Properties.Contains("maxPwdAge"))
                    {
                        maxPwdAge = TimeSpan.FromTicks((long)ouResult.Properties["maxPwdAge"][0]).Days * -1;
                        return maxPwdAge;
                    }
                    return maxPwdAge;
                }
            }
        }
        #endregion
    }
}

﻿using Newtonsoft.Json;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;


namespace EH.System.Commons
{
    public static class ADHelper
    {
        #region //   

        #region Server and Admin Info
        // LDAP地址 例如：LDAP://my.com.cn
        private const string LDAP_HOST = "LDAP://RootDSE";
        // 具有LDAP管理权限的特殊帐号
        private const string AdminUSER = "paceyl";
        // 具有LDAP管理权限的特殊帐号的密码
        private const string AdminPASSWORD = "Lpy@755815231";

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
        public static string GetDefaultNamingContext(string host = LDAP_HOST)
        {
            using (DirectoryEntry objRootDSE = new DirectoryEntry(host))
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
        public static bool CheckLogin(string userName, string password)
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
        public static void AddGroupMember(string groupName, string userName)
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
        public static void RemoveGroupMember(string groupName, string userName)
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
        public static object GetUserDNByName(string name)
        {
            var domain = GetDefaultNamingContext();

            DirectoryEntry entry = new DirectoryEntry(domain);
            DirectorySearcher search = new DirectorySearcher(entry);

            search.Filter = "(SAMAccountName=" + name + ")";

            SearchResult user = search.FindOne();

            if (user == null)
            {
                throw new Exception("请确认域用户是否正确");
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
         * 获取指定域组的信息
         * name 组名称 
         **/
        public static DirectoryEntry GetGroupByName(string name)
        {
            DirectorySearcher search = new DirectorySearcher(LDAP_HOST);
            search.SearchRoot = new DirectoryEntry(LDAP_HOST, AdminUSER, AdminPASSWORD);
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
        public static string ResetPassword(string userName)
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


        public static SearchResultCollection GetAllUser()
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
                    foreach (SearchResult result in results)
                    {
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

                                    return userResults;
                                }
                            }

                        }
                    }
                    return userResults;
                }
            }
        }

        #endregion




        //#region new//

        //#region Private Variables
        //private static string ADPath = "LDAP:/";//ConfigurationSettings.AppSettings["AD:Path"].ToString()/RootDSE;
        //private static string ADUser = "paceyl";//ConfigurationSettings.AppSettings["AD:AdminUser"].ToString();
        //private static string ADPassword = "Lpy@755815231";//ConfigurationSettings.AppSettings["AD:AdminPwd"].ToString();
        //private static string ADServer = "ehi.ehealth.com";
        //#endregion

        //#region Enumerations
        //public enum ADAccountOptions
        //{
        //    UF_TEMP_DUPLICATE_ACCOUNT = 0x0100,
        //    UF_NORMAL_ACCOUNT = 0x0200,
        //    UF_INTERDOMAIN_TRUST_ACCOUNT = 0x0800,
        //    UF_WORKSTATION_TRUST_ACCOUNT = 0x1000,
        //    UF_SERVER_TRUST_ACCOUNT = 0x2000,
        //    UF_DONT_EXPIRE_PASSWD = 0x10000,
        //    UF_SCRIPT = 0x0001,
        //    UF_ACCOUNTDISABLE = 0x0002,
        //    UF_HOMEDIR_REQUIRED = 0x0008,
        //    UF_LOCKOUT = 0x0010,
        //    UF_PASSWD_NOTREQD = 0x0020,
        //    UF_PASSWD_CANT_CHANGE = 0x0040,
        //    UF_ACCOUNT_LOCKOUT = 0X0010,
        //    UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = 0X0080,
        //}


        //public enum LoginResult
        //{
        //    LOGIN_OK = 0,
        //    LOGIN_USER_DOESNT_EXIST,
        //    LOGIN_USER_ACCOUNT_INACTIVE
        //}

        //#endregion

        //#region Methods

        ///// <summary>
        ///// This is used mainy for the logon process to ensure that the username and password match
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="Password"></param>
        ///// <returns></returns>
        //public static DirectoryEntry UserExists(string UserName, string Password)
        //{
        //    //create an instance of the DirectoryEntry
        //    DirectoryEntry de = GetDirectoryObject();//UserName,Password);

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    //set the search filter
        //    deSearch.SearchRoot = de;
        //    deSearch.Filter = "((objectClass=user)(cn=" + UserName + ")(userPassword=" + Password + "))";
        //    deSearch.SearchScope = SearchScope.Subtree;

        //    //set the property to return
        //    //deSearch.PropertiesToLoad.Add("givenName");

        //    //find the first instance
        //    SearchResult results = deSearch.FindOne();


        //    //if the username and password do match, then this implies a valid login
        //    //if so then return the DirectoryEntry object
        //    de = new DirectoryEntry(results.Path, ADUser, ADPassword, AuthenticationTypes.Secure);

        //    return de;

        //}


        //public static bool UserExists(string UserName)
        //{
        //    //create an instance of the DirectoryEntry
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    //set the search filter
        //    deSearch.SearchRoot = de;
        //    deSearch.Filter = "(&(objectClass=user) (cn=" + UserName + "))";

        //    //find the first instance
        //    SearchResultCollection results = deSearch.FindAll();

        //    //if the username and password do match, then this implies a valid login
        //    //if so then return the DirectoryEntry object
        //    if (results.Count == 0)
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }

        //}

        ///// <summary>
        ///// This method will not actually log a user in, but will perform tests to ensure
        ///// that the user account exists (matched by both the username and password), and also
        ///// checks if the account is active.
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="Password"></param>
        ///// <returns></returns>
        //public static ADHelper.LoginResult Login(string UserName, string Password)
        //{
        //    //first, check if the logon exists based on the username and password
        //    //DirectoryEntry de = GetUser(UserName,Password);

        //    if (IsUserValid(UserName, Password))
        //    {
        //        DirectoryEntry de = GetUser(UserName);
        //        if (de != null)
        //        {
        //            //convert the accountControl value so that a logical operation can be performed
        //            //to check of the Disabled option exists.
        //            int userAccountControl = Convert.ToInt32(de.Properties["userAccountControl"][0]);
        //            de.Close();

        //            //if the disabled item does not exist then the account is active
        //            if (!IsAccountActive(userAccountControl))
        //            {
        //                return LoginResult.LOGIN_USER_ACCOUNT_INACTIVE;
        //            }
        //            else
        //            {
        //                return LoginResult.LOGIN_OK;
        //            }

        //        }
        //        else
        //        {
        //            return LoginResult.LOGIN_USER_DOESNT_EXIST;
        //        }
        //    }
        //    else
        //    {
        //        return LoginResult.LOGIN_USER_DOESNT_EXIST;
        //    }
        //}

        ///// <summary>
        ///// This will perfrom a logical operation on the userAccountControl values
        ///// to see if the user account is enabled or disabled.  The flag for determining if the
        ///// account is active is a bitwise value (decimal =2)
        ///// </summary>
        ///// <param name="userAccountControl"></param>
        ///// <returns></returns>
        //public static bool IsAccountActive(int userAccountControl)
        //{
        //    int userAccountControl_Disabled = Convert.ToInt32(ADAccountOptions.UF_ACCOUNTDISABLE);
        //    int flagExists = userAccountControl & userAccountControl_Disabled;
        //    //if a match is found, then the disabled flag exists within the control flags
        //    if (flagExists > 0)
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}
        ///// <summary>
        ///// This method will attempt to log in a user based on the username and password
        ///// to ensure that they have been set up within the Active Directory.  This is the basic UserName, Password
        ///// check.
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="Password"></param>
        ///// <returns></returns>
        //public static bool IsUserValid(string UserName, string Password)
        //{
        //    try
        //    {
        //        //if the object can be created then return true
        //        DirectoryEntry deUser = GetUser(UserName, Password);
        //        deUser.Close();
        //        return true;
        //    }
        //    catch (Exception)
        //    {
        //        //otherwise return false
        //        return false;
        //    }
        //}

        //#region Search Methods
        ///// <summary>
        ///// This will return a DirectoryEntry object if the user does exist
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <returns></returns>
        //public static DirectoryEntry GetUser(string UserName)
        //{
        //    //create an instance of the DirectoryEntry
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    deSearch.SearchRoot = de;
        //    //set the search filter
        //    deSearch.Filter = "(&(objectClass=user)(cn=" + UserName + "))";
        //    deSearch.SearchScope = SearchScope.Subtree;

        //    //find the first instance
        //    SearchResult results = deSearch.FindOne();

        //    //if found then return, otherwise return Null
        //    if (results != null)
        //    {
        //        de = new DirectoryEntry(results.Path, ADUser, ADPassword, AuthenticationTypes.Secure);
        //        //if so then return the DirectoryEntry object
        //        return de;
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        ///// <summary>
        ///// Override method which will perfrom query based on combination of username and password
        ///// This is used with the login process to validate the user credentials and return a user
        ///// object for further validation.  This is slightly different from the other GetUser... methods as this
        ///// will use the UserName and Password supplied as the authentication to check if the user exists, if so then
        ///// the users object will be queried using these credentials.s
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="password"></param>
        ///// <returns></returns>
        //public static DirectoryEntry GetUser(string UserName, string Password)
        //{
        //    //create an instance of the DirectoryEntry
        //    DirectoryEntry de = GetDirectoryObject(UserName, Password);

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    deSearch.SearchRoot = de;
        //    //set the search filter
        //    deSearch.Filter = "(&(objectClass=user)(cn=" + UserName + "))";
        //    deSearch.SearchScope = SearchScope.Subtree;

        //    //set the property to return
        //    //deSearch.PropertiesToLoad.Add("givenName");

        //    //find the first instance
        //    SearchResult results = deSearch.FindOne();

        //    //if a match is found, then create directiry object and return, otherwise return Null
        //    if (results != null)
        //    {
        //        //create the user object based on the admin priv.
        //        de = new DirectoryEntry(results.Path, ADUser, ADPassword, AuthenticationTypes.Secure);
        //        return de;
        //    }
        //    else
        //    {
        //        return null;
        //    }


        //}
        ///// <summary>
        ///// This will take a username and query the AD for the user.  When found it will transform
        ///// the results from the poperty collection into a Dataset which can be used by the client
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <returns></returns>
        //public static DataSet GetUserDataSet(string UserName)
        //{
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    deSearch.SearchRoot = de;
        //    //set the search filter
        //    deSearch.Filter = "(&(objectClass=user)(cn=" + UserName + "))";
        //    deSearch.SearchScope = SearchScope.Subtree;

        //    //find the first instance
        //    SearchResult results = deSearch.FindOne();

        //    //get Empty user dataset
        //    DataSet dsUser = CreateUserDataSet();

        //    //If no user record returned, then dont do anything, otherwise
        //    //populate
        //    if (results != null)
        //    {
        //        //populate the dataset with the values from the results
        //        dsUser.Tables["User"].Rows.Add(PopulateUserDataSet(results, dsUser.Tables["User"]));

        //    }
        //    de.Close();

        //    return dsUser;

        //}

        ///// <summary>
        ///// This method will return a dataset of user details based on criteria
        ///// passed to the query.  The criteria is in the LDAP format ie
        ///// (cn='xxx')(sn='eee') etc
        ///// </summary>
        ///// <param name="Criteria"></param>
        ///// <returns></returns>
        //public static DataSet GetUsersDataSet(string Criteria)
        //{
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    deSearch.SearchRoot = de;
        //    //set the search filter
        //    deSearch.Filter = "(&(objectClass=user)(objectCategory=person)" + Criteria + ")";
        //    deSearch.SearchScope = SearchScope.Subtree;

        //    //find the first instance
        //    SearchResultCollection results = deSearch.FindAll();

        //    //get Empty user dataset
        //    DataSet dsUser = CreateUserDataSet();

        //    //If no user record returned, then dont do anything, otherwise
        //    //populate
        //    if (results.Count > 0)
        //    {
        //        foreach (SearchResult result in results)
        //        {
        //            //populate the dataset with the values from the results
        //            dsUser.Tables["User"].Rows.Add(PopulateUserDataSet(result, dsUser.Tables["User"]));
        //        }
        //    }

        //    de.Close();
        //    return dsUser;

        //}

        ///// <summary>
        ///// This method will query all of the defined AD groups
        ///// and will turn the results into a dataset to be returned
        ///// </summary>
        ///// <returns></returns>
        //public static DataSet GetGroups()
        //{
        //    DataSet dsGroup = new DataSet();
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    //set the search filter
        //    deSearch.SearchRoot = de;
        //    //deSearch.PropertiesToLoad.Add("cn");
        //    deSearch.Filter = "(&(objectClass=group)(cn=CS_*))";

        //    //find the first instance
        //    SearchResultCollection results = deSearch.FindAll();

        //    //Create a new table object within the dataset
        //    DataTable tbGroup = dsGroup.Tables.Add("Groups");
        //    tbGroup.Columns.Add("GroupName");

        //    //if there are results (there should be some!!), then convert the results
        //    //into a dataset to be returned.
        //    if (results.Count > 0)
        //    {

        //        //iterate through collection and populate the table with
        //        //the Group Name
        //        foreach (SearchResult Result in results)
        //        {
        //            //set a new empty row
        //            DataRow rwGroup = tbGroup.NewRow();

        //            //populate the column
        //            rwGroup["GroupName"] = Result.Properties["cn"][0];

        //            //append the row to the table of the dataset
        //            tbGroup.Rows.Add(rwGroup);
        //        }
        //    }
        //    return dsGroup;
        //}

        ///// <summary>
        ///// This method will return all users for the specified group in a dataset
        ///// </summary>
        ///// <param name="GroupName"></param>
        ///// <returns></returns>
        //public static DataSet GetUsersForGroup(string GroupName)
        //{
        //    DataSet dsUser = new DataSet();
        //    DirectoryEntry de = GetDirectoryObject();

        //    //create instance fo the direcory searcher
        //    DirectorySearcher deSearch = new DirectorySearcher();

        //    //set the search filter
        //    deSearch.SearchRoot = de;
        //    //deSearch.PropertiesToLoad.Add("cn");
        //    deSearch.Filter = "(&(objectClass=group)(cn=" + GroupName + "))";

        //    //get the group result
        //    SearchResult results = deSearch.FindOne();

        //    //Create a new table object within the dataset
        //    DataTable tbUser = dsUser.Tables.Add("Users");
        //    tbUser.Columns.Add("UserName");
        //    tbUser.Columns.Add("DisplayName");
        //    tbUser.Columns.Add("EMailAddress");

        //    //Create default row
        //    DataRow rwDefaultUser = tbUser.NewRow();
        //    rwDefaultUser["UserName"] = "0";
        //    rwDefaultUser["DisplayName"] = "(Not Specified)";
        //    rwDefaultUser["EMailAddress"] = "(Not Specified)";
        //    tbUser.Rows.Add(rwDefaultUser);

        //    //if the group is valid, then continue, otherwise return a blank dataset
        //    if (results != null)
        //    {
        //        //create a link to the group object, so we can get the list of members
        //        //within the group
        //        DirectoryEntry deGroup = new DirectoryEntry(results.Path, ADUser, ADPassword, AuthenticationTypes.Secure);
        //        //assign a property collection
        //        System.DirectoryServices.PropertyCollection pcoll = deGroup.Properties;
        //        int n = pcoll["member"].Count;

        //        //if there are members fo the group, then get the details and assign to the table
        //        for (int l = 0; l < n; l++)
        //        {
        //            //create a link to the user object sot hat the FirstName, LastName and SUername can be gotten
        //            DirectoryEntry deUser = new DirectoryEntry(ADPath + "/" + pcoll["member"][l].ToString(), ADUser, ADPassword, AuthenticationTypes.Secure);

        //            //set a new empty row
        //            DataRow rwUser = tbUser.NewRow();

        //            //populate the column
        //            rwUser["UserName"] = GetProperty(deUser, "cn");
        //            rwUser["DisplayName"] = GetProperty(deUser, "givenName") + " " + GetProperty(deUser, "sn");
        //            rwUser["EMailAddress"] = GetProperty(deUser, "mail");
        //            //append the row to the table of the dataset
        //            tbUser.Rows.Add(rwUser);

        //            //close the directory entry object
        //            deUser.Close();

        //        }
        //        de.Close();
        //        deGroup.Close();
        //    }


        //    return dsUser;
        //}

        //#endregion

        ///// <summary>
        ///// This will query the user (by using the administrator role) and will set the new password
        ///// This will not validate the existing password, as it will be assumed that if there logged in then
        ///// the password can be changed.
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="OldPassword"></param>
        ///// <param name="NewPassword"></param>
        //public static void SetUserPassword(string UserName, string NewPassword)
        //{
        //    //get reference to user
        //    string LDAPDomain = "/CN=" + UserName + ",CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oUser = GetDirectoryObject(LDAPDomain);//,UserName,OldPassword);
        //    oUser.Invoke("SetPassword", new Object[] { NewPassword });
        //    oUser.Close();
        //}

        ///// <summary>
        ///// This method will be used by the admin query screen, and is a method
        ///// to return users based on a possible combination of lastname, email address or corporate
        ///// </summary>
        ///// <param name="Lastname"></param>
        ///// <param name="EmailAddress"></param>
        ///// <param name="Corporate"></param>
        ///// <returns></returns>
        //public static DataSet GetUsersByNameEmailCorporate(string LastName, string EmailAddress, string Corporate)
        //{
        //    StringBuilder SQLWhere = new StringBuilder();

        //    //if the LastName is present, then include in the where clause
        //    if (LastName != string.Empty)
        //    {
        //        SQLWhere.Append("(sn=" + LastName + ")");
        //    }


        //    //if the emailaddress is present, then include in the where clause
        //    if (EmailAddress != string.Empty)
        //    {
        //        SQLWhere.Append("(mail=" + EmailAddress + ")");
        //    }

        //    //if the corporate is present, then include in the where clause
        //    if ((Corporate != string.Empty) && (Corporate != "1"))
        //    {
        //        SQLWhere.Append("(extensionAttribute12=" + Corporate + ")");
        //    }

        //    //append the where clause, remove the last 'AND'
        //    //SQLStmt.Append(";(objectClass=*); sn, givenname, mail");

        //    return GetUsersDataSet(SQLWhere.ToString());

        //}
        //#region Set User Details Methods
        ///// <summary>
        ///// Set the user password
        ///// </summary>
        ///// <param name="oDE"></param>
        ///// <param name="Password"></param>
        //public static void SetUserPassword(DirectoryEntry oDE, string Password)
        //{

        //    oDE.Invoke("SetPassword", new Object[] { Password });

        //    //string[] yourpw={Password};
        //    //oDE.Invoke("SetPassword", yourpw);
        //    //oDE.CommitChanges();

        //    //object[] password = new object[] {Password};
        //    //object ret = oDE.Invoke("SetPassword", password );
        //    //oDE.CommitChanges();

        //}
        ///// <summary>
        ///// This will enable a user account based on the username
        ///// </summary>
        ///// <param name="UserName"></param>
        //public static void EnableUserAccount(string UserName)
        //{
        //    //get the directory entry fot eh user and enable the password
        //    EnableUserAccount(GetUser(UserName));
        //}

        //public static void EnableUserAccount(DirectoryEntry oDE)
        //{
        //    //we enable the account by resetting all the account options excluding the disable flag
        //    oDE.Properties["userAccountControl"][0] = ADHelper.ADAccountOptions.UF_NORMAL_ACCOUNT | ADHelper.ADAccountOptions.UF_DONT_EXPIRE_PASSWD;
        //    oDE.CommitChanges();

        //    //			oDE.Invoke("accountDisabled",new Object[]{"false"});
        //    oDE.Close();
        //}


        ///// <summary>
        ///// This will disable the user account based on the username passed to it
        ///// </summary>
        ///// <param name="Username"></param>
        //public static void DisableUserAccount(string UserName)
        //{
        //    //get the directory entry fot eh user and enable the password
        //    DisableUserAccount(GetUser(UserName));
        //}


        ///// <summary>
        ///// Enable the user account based on the DirectoryEntry object passed to it
        ///// </summary>
        ///// <param name="oDE"></param>
        //public static void DisableUserAccount(DirectoryEntry oDE)
        //{
        //    //we disable the account by resetting all the default properties
        //    oDE.Properties["userAccountControl"][0] = ADHelper.ADAccountOptions.UF_NORMAL_ACCOUNT | ADHelper.ADAccountOptions.UF_DONT_EXPIRE_PASSWD | ADHelper.ADAccountOptions.UF_ACCOUNTDISABLE;
        //    oDE.CommitChanges();
        //    //			oDE.Invoke("accountDisabled",new Object[]{"true"});
        //    oDE.Close();
        //}

        ///// <summary>
        ///// Override method for adding a user to a group.  The group will be specified
        ///// so that a group object can be located, then the user will be queried and added to the group
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="GroupName"></param>
        //public static void AddUserToGroup(string UserName, string GroupName)
        //{
        //    string LDAPDomain = string.Empty;
        //    //get reference to group
        //    LDAPDomain = "/CN=" + GroupName + ",CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oGroup = GetDirectoryObject(LDAPDomain);

        //    //get reference to user
        //    LDAPDomain = "/CN=" + UserName + ",CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oUser = GetDirectoryObject(LDAPDomain);

        //    //Add the user to the group via the invoke method
        //    oGroup.Invoke("Add", new Object[] { oUser.Path.ToString() });

        //    oGroup.Close();
        //    oUser.Close();
        //}

        //public static void RemoveUserFromGroup(string UserName, string GroupName)
        //{
        //    string LDAPDomain = string.Empty;

        //    //get reference to group
        //    LDAPDomain = "/CN=" + GroupName + ",CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oGroup = GetDirectoryObject(LDAPDomain);

        //    //get reference to user
        //    LDAPDomain = "/CN=" + UserName + ",CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oUser = GetDirectoryObject(LDAPDomain);

        //    //Add the user to the group via the invoke method
        //    oGroup.Invoke("Remove", new Object[] { oUser.Path.ToString() });

        //    oGroup.Close();
        //    oUser.Close();
        //}


        //#endregion

        //#region Helper Methods
        ///// <summary>
        ///// This will retreive the specified poperty value from the DirectoryEntry object (if the property exists)
        ///// </summary>
        ///// <param name="oDE"></param>
        ///// <param name="PropertyName"></param>
        ///// <returns></returns>
        //public static string GetProperty(DirectoryEntry oDE, string PropertyName)
        //{
        //    if (oDE.Properties.Contains(PropertyName))
        //    {
        //        return oDE.Properties[PropertyName][0].ToString();
        //    }
        //    else
        //    {
        //        return string.Empty;
        //    }
        //}

        ///// <summary>
        ///// This is an override that will allow a property to be extracted directly from
        ///// a searchresult object
        ///// </summary>
        ///// <param name="searchResult"></param>
        ///// <param name="PropertyName"></param>
        ///// <returns></returns>
        //public static string GetProperty(SearchResult searchResult, string PropertyName)
        //{
        //    if (searchResult.Properties.Contains(PropertyName))
        //    {
        //        return searchResult.Properties[PropertyName][0].ToString();
        //    }
        //    else
        //    {
        //        return string.Empty;
        //    }
        //}
        ///// <summary>
        ///// This will test the value of the propertyvalue and if empty will not set the property
        ///// as AD is particular about being sent blank values
        ///// </summary>
        ///// <param name="oDE"></param>
        ///// <param name="PropertyName"></param>
        ///// <param name="PropertyValue"></param>
        //public static void SetProperty(DirectoryEntry oDE, string PropertyName, string PropertyValue)
        //{
        //    //check if the value is valid, otherwise dont update
        //    if (PropertyValue != string.Empty)
        //    {
        //        //check if the property exists before adding it to the list
        //        if (oDE.Properties.Contains(PropertyName))
        //        {
        //            oDE.Properties[PropertyName][0] = PropertyValue;
        //        }
        //        else
        //        {
        //            oDE.Properties[PropertyName].Add(PropertyValue);
        //        }
        //    }
        //}

        ///// <summary>
        ///// This is an internal method for retreiving a new directoryentry object
        ///// </summary>
        ///// <returns></returns>
        //private static DirectoryEntry GetDirectoryObject()
        //{
        //    DirectoryEntry oDE;

        //    oDE = new DirectoryEntry(ADPath, ADUser, ADPassword, AuthenticationTypes.Secure);

        //    return oDE;
        //}

        ///// <summary>
        ///// Override function that that will attempt a logon based on the users credentials
        ///// </summary>
        ///// <param name="UserName"></param>
        ///// <param name="Password"></param>
        ///// <returns></returns>
        //private static DirectoryEntry GetDirectoryObject(string UserName, string Password)
        //{
        //    DirectoryEntry oDE;

        //    oDE = new DirectoryEntry(ADPath, UserName, Password, AuthenticationTypes.Secure);

        //    return oDE;
        //}

        ///// <summary>
        ///// This will create the directory entry based on the domain object to return
        ///// The DomainReference will contain the qualified syntax for returning an entry
        ///// at the location rather than returning the root.  
        ///// i.e. /CN=Users,DC=creditsights, DC=cyberelves, DC=Com
        ///// </summary>
        ///// <param name="DomainReference"></param>
        ///// <returns></returns>
        //private static DirectoryEntry GetDirectoryObject(string DomainReference)
        //{
        //    DirectoryEntry oDE;

        //    oDE = new DirectoryEntry(ADPath + DomainReference, ADUser, ADPassword, AuthenticationTypes.Secure);

        //    return oDE;
        //}

        ///// <summary>
        ///// Addition override that will allow ovject to be created based on the users credentials.
        ///// This is useful for instances such as setting password etc.
        ///// </summary>
        ///// <param name="DomainReference"></param>
        ///// <param name="UserName"></param>
        ///// <param name="Password"></param>
        ///// <returns></returns>
        //private static DirectoryEntry GetDirectoryObject(string DomainReference, string UserName, string Password)
        //{
        //    DirectoryEntry oDE;

        //    oDE = new DirectoryEntry(ADPath + DomainReference, UserName, Password, AuthenticationTypes.Secure);

        //    return oDE;
        //}


        //#endregion

        //#region Internal Methods
        ///// <summary>
        ///// This method will create a new directory object and pass it back so that
        ///// it can be populated
        ///// </summary>
        ///// <param name="cn"></param>
        ///// <returns></returns>
        //public static DirectoryEntry CreateNewUser(string cn)
        //{
        //    //set the LDAP qualification so that the user will be created under the Users
        //    //container
        //    string LDAPDomain = "/CN=Users," + GetLDAPDomain();
        //    DirectoryEntry oDE = GetDirectoryObject(LDAPDomain);
        //    DirectoryEntry oDEC = oDE.Children.Add("CN=" + cn, "User");
        //    oDE.Close();
        //    return oDEC;

        //}

        ///// <summary>
        ///// This will read in the ADServer value from the web.config and will return it
        ///// as an LDAP path ie DC=creditsights, DC=cyberelves, DC=com.
        ///// This is required when creating directoryentry other than the root.
        ///// </summary>
        ///// <returns></returns>
        //private static string GetLDAPDomain()
        //{
        //    StringBuilder LDAPDomain = new StringBuilder();
        //    string[] LDAPDC = ADServer.Split('.');//ConfigurationSettings.AppSettings["AD:Server"].Split('.');

        //    for (int i = 0; i < LDAPDC.GetUpperBound(0) + 1; i++)
        //    {
        //        LDAPDomain.Append("DC=" + LDAPDC[i]);
        //        if (i < LDAPDC.GetUpperBound(0))
        //        {
        //            LDAPDomain.Append(",");
        //        }
        //    }

        //    return LDAPDomain.ToString();
        //}


        ///// <summary>
        ///// This method will create a Dataset stucture containing all relevant fields
        ///// that match to a user.
        ///// </summary>
        ///// <returns></returns>
        //private static DataSet CreateUserDataSet()
        //{

        //    DataSet ds = new DataSet();
        //    //Create a new table object within the dataset
        //    DataTable tb = ds.Tables.Add("User");

        //    //Create all the columns
        //    tb.Columns.Add("LoginName");
        //    tb.Columns.Add("FirstName");
        //    tb.Columns.Add("MiddleInitial");
        //    tb.Columns.Add("LastName");
        //    tb.Columns.Add("Address1");
        //    tb.Columns.Add("Address2");
        //    tb.Columns.Add("Title");
        //    tb.Columns.Add("Company");
        //    tb.Columns.Add("City");
        //    tb.Columns.Add("State");
        //    tb.Columns.Add("Country");
        //    tb.Columns.Add("Zip");
        //    tb.Columns.Add("Phone");
        //    tb.Columns.Add("Extension");
        //    tb.Columns.Add("Fax");
        //    tb.Columns.Add("EmailAddress");
        //    tb.Columns.Add("ChallengeQuestion");
        //    tb.Columns.Add("ChallengeResponse");
        //    tb.Columns.Add("MemberCompany");
        //    tb.Columns.Add("CompanyRelationShipExists");
        //    tb.Columns.Add("Status");
        //    tb.Columns.Add("AssignedSalesPerson");
        //    tb.Columns.Add("AcceptTAndC");
        //    tb.Columns.Add("Jobs");
        //    tb.Columns.Add("Email_Overnight");
        //    tb.Columns.Add("Email_DailyEmergingMarkets");
        //    tb.Columns.Add("Email_DailyCorporateAlerts");
        //    tb.Columns.Add("AssetMgtRange");
        //    tb.Columns.Add("ReferralCompany");
        //    tb.Columns.Add("CorporateAffiliation");
        //    tb.Columns.Add("DateCreated");
        //    tb.Columns.Add("DateLastModified");
        //    tb.Columns.Add("DateOfExpiry");
        //    tb.Columns.Add("AccountIsActive");

        //    return ds;
        //}

        ///// <summary>
        ///// This method will return a DataRow object which will be added to the userdataset object
        ///// This will also allow the iteration of multiple rows
        ///// </summary>
        ///// <param name="userSearchResult"></param>
        ///// <returns></returns>
        //private static DataRow PopulateUserDataSet(SearchResult userSearchResult, DataTable userTable)
        //{
        //    //set a new empty row
        //    DataRow rwUser = userTable.NewRow();

        //    rwUser["LoginName"] = GetProperty(userSearchResult, "cn");
        //    rwUser["FirstName"] = GetProperty(userSearchResult, "givenName");
        //    rwUser["MiddleInitial"] = GetProperty(userSearchResult, "initials");
        //    rwUser["LastName"] = GetProperty(userSearchResult, "sn");

        //    string tempAddress = GetProperty(userSearchResult, "homePostalAddress");
        //    //if the address does not exist, then default to blank fields
        //    if (tempAddress != string.Empty)
        //    {
        //        string[] addressArray = tempAddress.Split(';');
        //        rwUser["Address1"] = addressArray[0];
        //        rwUser["Address2"] = addressArray[1];
        //    }
        //    else
        //    {
        //        rwUser["Address1"] = string.Empty;
        //        rwUser["Address2"] = string.Empty;
        //    }

        //    rwUser["Title"] = GetProperty(userSearchResult, "title");
        //    rwUser["Company"] = GetProperty(userSearchResult, "company");
        //    rwUser["State"] = GetProperty(userSearchResult, "st");
        //    rwUser["City"] = GetProperty(userSearchResult, "l");
        //    rwUser["Country"] = GetProperty(userSearchResult, "co");
        //    rwUser["Zip"] = GetProperty(userSearchResult, "postalCode");
        //    rwUser["Phone"] = GetProperty(userSearchResult, "telephoneNumber");
        //    rwUser["Extension"] = GetProperty(userSearchResult, "otherTelephone");
        //    rwUser["Fax"] = GetProperty(userSearchResult, "facsimileTelephoneNumber");
        //    rwUser["EmailAddress"] = GetProperty(userSearchResult, "mail");
        //    rwUser["ChallengeQuestion"] = GetProperty(userSearchResult, "extensionAttribute1");
        //    rwUser["ChallengeResponse"] = GetProperty(userSearchResult, "extensionAttribute2");
        //    rwUser["MemberCompany"] = GetProperty(userSearchResult, "extensionAttribute3");
        //    rwUser["CompanyRelationShipExists"] = GetProperty(userSearchResult, "extensionAttribute4");
        //    rwUser["Status"] = GetProperty(userSearchResult, "extensionAttribute5");
        //    rwUser["AssignedSalesPerson"] = GetProperty(userSearchResult, "extensionAttribute6");
        //    rwUser["AcceptTAndC"] = GetProperty(userSearchResult, "extensionAttribute7");
        //    rwUser["Jobs"] = GetProperty(userSearchResult, "extensionAttribute8");

        //    //handle the split of the email options
        //    string tempTempEmail = GetProperty(userSearchResult, "extensionAttribute9");

        //    //if no email address are present, then default to blank
        //    if (tempTempEmail != string.Empty)
        //    {
        //        string[] emailArray = tempTempEmail.Split(';');
        //        rwUser["Email_Overnight"] = emailArray[0];
        //        rwUser["Email_DailyEmergingMarkets"] = emailArray[1];
        //        rwUser["Email_DailyCorporateAlerts"] = emailArray[2];
        //    }
        //    else
        //    {
        //        rwUser["Email_Overnight"] = "false";
        //        rwUser["Email_DailyEmergingMarkets"] = "false";
        //        rwUser["Email_DailyCorporateAlerts"] = "false";
        //    }

        //    rwUser["AssetMgtRange"] = GetProperty(userSearchResult, "extensionAttribute10");
        //    rwUser["ReferralCompany"] = GetProperty(userSearchResult, "extensionAttribute11");
        //    rwUser["CorporateAffiliation"] = GetProperty(userSearchResult, "extensionAttribute12");
        //    rwUser["DateCreated"] = GetProperty(userSearchResult, "whenCreated");
        //    rwUser["DateLastModified"] = GetProperty(userSearchResult, "whenChanged");
        //    rwUser["DateOfExpiry"] = GetProperty(userSearchResult, "extensionAttribute12");
        //    rwUser["AccountIsActive"] = IsAccountActive(Convert.ToInt32(GetProperty(userSearchResult, "userAccountControl")));
        //    return rwUser;

        //}
        //#endregion



        //#endregion
        //#endregion

    }
}

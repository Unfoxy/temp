using System;
using System.Linq;
using System.Net.NetworkInformation;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.ServiceProcess;
using System.Management;
using Microsoft.Win32;
using System.Windows.Forms.VisualStyles;

namespace desktopDashboard___Y_Lee
{
    public class Functions
    {
        public static Tuple<string[], int> disposeMemberships(string[] disposeMemberships, int count, string addDispose)
        {
            disposeMemberships[count] = addDispose;
            count++;


            return Tuple.Create(disposeMemberships, count);
        }
        public static Tuple<string[], int> removeMembership(string[] currentMemberships, int count, string disposeMembership)
        {
            currentMemberships = currentMemberships.Where(w => w != disposeMembership).ToArray();
            count--;

            //for (int i = 0; i < count; i++)
            //{
            //    if (currentMemberships[i] == disposeMembership)
            //        currentMemberships[i] = null;
            //}
            //count--;
            return Tuple.Create(currentMemberships, count);
        }

        public static Tuple<string[], int> displayMembership(string disposeMembership, string user)
        {
            PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "in1.ad.innovene.com", "OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com");
            UserPrincipal sourceUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, user);


            if (sourceUser != null)
            {
                var sourceGroups = sourceUser.GetGroups();

                int count = 0;
                string[] membership = new string[1000];
                foreach (Principal sourceGroup in sourceGroups)
                {
                    if (sourceGroup.Name != disposeMembership)
                    {
                        membership[count] = sourceGroup.Name;
                        count++;
                    }
                    else
                    {
                        count++;
                    }
                }
                return Tuple.Create(membership, count);
            }
            else
                return null;
        }
        //public static string[] displayMembership(string disposeMembership, string user)
        //{
        //    PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "in1.ad.innovene.com", "OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com");
        //    UserPrincipal sourceUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, user);


        //    if (sourceUser != null)
        //    {
        //        var sourceGroups = sourceUser.GetGroups();

        //        int count = 0;
        //        string[] membership = new string[1000];
        //        foreach (Principal sourceGroup in sourceGroups)
        //        {
        //            if (sourceGroup.Name != disposeMembership)
        //            {
        //                membership[count] = sourceGroup.Name;
        //                count++;
        //            }
        //            else
        //            {
        //                count++;
        //            }    
        //        }
        //        return membership;
        //    }
        //    else
        //        return null;
        //}
        public static void resetPassword(string username)
        {
            //PrincipalContext context = new PrincipalContext(ContextType.Domain);                                      //It works either way. Leaving here to be referenced for createUser.
            //UserPrincipal user = UserPrincipal.FindByIdentity(context, IdentityType.SamAccountName, username);
            ////user.Enabled = true;            //Enable Account if it is disabled - Not Working right now
            //user.SetPassword("Ineos2023");
            //user.ExpirePasswordNow();         //Force user to change password at next logon
            //user.Save();

            DirectoryEntry account = new DirectoryEntry("LDAP://OU=Client,DC=in1,DC=ad,DC=innovene,DC=com", null, null, AuthenticationTypes.Secure | AuthenticationTypes.Sealing | AuthenticationTypes.Signing);
            DirectorySearcher search = new DirectorySearcher(account);
            search.Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))";
            account = search.FindOne().GetDirectoryEntry();

            account.Invoke("SetPassword", "Ineos2023");
            account.Properties["LockOutTime"].Value = 0;        //Unlock Account
            account.Properties["pwdLastSet"][0] = 0;            //Prompt User to Reset Password
            account.CommitChanges();
        }
        public static string[] GetAD(string username)
        {
            try
            {
                DirectoryEntry ldapConnection = new DirectoryEntry("");
                ldapConnection.Path = "LDAP://OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";           //For all Ineos sites
                ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
                DirectorySearcher search = new DirectorySearcher(ldapConnection);

                search.Filter = "(|"                                                                //Multiple search options
                                   + "(sAMAccountName=" + username + ")"
                                   + "(mail=" + username + ")"
                                   + "(mail=" + username + "@ineos.com" + ")"
                                   + "(cn=" + username + ")"
                                   + ")";

                string[] requiredProperties = new string[] { "cn",                                  //Name, mail, NTID, Site, Ofiice, Telephone, Job title
                                                           "mail",
                                                 "sAMAccountName",
                                            "extensionAttribute1",
                                     "physicalDeliveryOfficeName",
                                                "telephoneNumber",
                                                          "title" };

                foreach (String property in requiredProperties)
                    search.PropertiesToLoad.Add(property);

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    string[] results = new string[7] { "", "", "", "", "", "", "" };
                    int i = 0;
                    foreach (String property in requiredProperties)
                        foreach (Object myCollection in result.Properties[property])
                        {
                            if (results[i] != "")
                                i++;
                            results[i] = myCollection.ToString();
                        }
                    return results;
                }
                else
                    return null;
            }
            catch
            {
                return null;
            }
        }
        public static Tuple<string[], string[], string[], int> queryAD(string site, string filter)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry("");
            ldapConnection.Path = "LDAP://OU=Client,DC=in1,DC=ad,DC=innovene,DC=com"; //Default All Sites
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            DirectorySearcher search = new DirectorySearcher(ldapConnection);
            var expiredTime = DateTime.Now.AddDays(-180).ToFileTime();                //Account expiration date 180days

            if (site.ToUpper() == "BMC")
                ldapConnection.Path = "LDAP://OU=BMC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "MVW")
                ldapConnection.Path = "LDAP://OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "CHO")
                ldapConnection.Path = "LDAP://OU=CHO,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "LAR")
                ldapConnection.Path = "LDAP://OU=LAR,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "HDC")
                ldapConnection.Path = "LDAP://OU=HDC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";

            if (filter == "Account Deactivated Users")
            {
                search.Filter = "(&"
                                    + "(userAccountControl=" + "514" + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + ")";
            }
            else if (filter == "Password Expired Users")
            {
                search.Filter = "(&"
                                    + "(pwdLastSet<=" + expiredTime + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + "(userAccountControl=" + "512" + ")"
                                    + ")";
            }
            else if (filter == "Account Locked Users")
            {
                search.Filter = "(&"
                                    + "(lockoutTime>=" + "1" + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + ")";
            }
            else if (filter == "Account Expired Users")
            {
                search.Filter = "(&"
                                   + "(accountExpires<=" + expiredTime + ")"
                                   + "(userAccountControl=" + "512" + ")"
                                   + "(extensionAttribute2=" + "OP USA" + ")"
                                   + "(!" + "(accountExpires=" + "0" + ")"
                                   + ")"
                                   + ")";
            }

            try
            {
                string[] requiredProperties = new string[] { "cn", "sAMAccountName", "mail" };      //Name, NTID, Email from AD Attributes

                foreach (String property in requiredProperties)
                    search.PropertiesToLoad.Add(property);

                SearchResultCollection result = search.FindAll();

                string[] name = new string[1000];       //Max number for the output
                string[] ntid = new string[1000];       //Max number for the output
                string[] email = new string[1000];      //Max number for the output
                int count = 0;
                foreach (SearchResult userResults in result)
                {
                    if (!userResults.Properties.Contains("cn"))
                        name[count] = "";
                    else
                        name[count] = userResults.Properties["cn"][0].ToString();

                    if (!userResults.Properties.Contains("sAMAccountName"))
                        ntid[count] = "";
                    else
                        ntid[count] = userResults.Properties["sAMAccountName"][0].ToString();

                    if (!userResults.Properties.Contains("mail"))
                        email[count] = "";
                    else
                        email[count] = userResults.Properties["mail"][0].ToString();

                    count++;
                }
                return Tuple.Create(name, ntid, email, count);
            }
            catch
            {
                throw;
            }
        }
        public static int userCount(string site, string filter)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry("");
            ldapConnection.Path = "LDAP://OU=Client,DC=in1,DC=ad,DC=innovene,DC=com"; //Default All Sites
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            DirectorySearcher search = new DirectorySearcher(ldapConnection);
            var expiredTime = DateTime.Now.AddDays(-180).ToFileTime();                //Account expiration date 180days

            if (site.ToUpper() == "BMC:")
                ldapConnection.Path = "LDAP://OU=BMC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "MVW:")
                ldapConnection.Path = "LDAP://OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "CHO:")
                ldapConnection.Path = "LDAP://OU=CHO,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "LAR:")
                ldapConnection.Path = "LDAP://OU=LAR,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (site.ToUpper() == "HDC:")
                ldapConnection.Path = "LDAP://OU=HDC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";

            if (filter == "Account Deactivated Users")
            {
                search.Filter = "(&"
                                    + "(userAccountControl=" + "514" + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + ")";
            }
            else if (filter == "Password Expired Users")
            {
                search.Filter = "(&"
                                    + "(pwdLastSet<=" + expiredTime + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + "(userAccountControl=" + "512" + ")"
                                    + ")";
            }
            else if (filter == "Account Locked Users")
            {
                search.Filter = "(&"
                                    + "(lockoutTime>=" + "1" + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + ")";
            }
            else if (filter == "Account Expired Users")
            {
                search.Filter = "(&"
                                    + "(accountExpires<=" + expiredTime + ")"
                                    + "(userAccountControl=" + "512" + ")"
                                    + "(extensionAttribute2=" + "OP USA" + ")"
                                    + "(!" + "(accountExpires=" + "0" + ")"
                                    + ")"
                                    + ")";
            }

            try
            {
                SearchResultCollection result = search.FindAll();
                return result.Count;
            }
            catch
            {
                throw;
            }
        }
        public static void editRegistry(string hostname)
        {
            var inputRegistry = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, hostname);
            inputRegistry = inputRegistry.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\0001", true); //IPChecksumOffloadIPv4 location
            inputRegistry.SetValue("*IPChecksumOffloadIPv4", "0");
        }
        public static void startupChange(string hostname)
        {
            ManagementBaseObject inParam;
            ManagementBaseObject outParam;
            var serviceController = new ServiceController();
            ManagementObject obj = new ManagementObject(@"\\" + hostname + "\\root\\cimv2:Win32_Service.Name='RemoteRegistry'");
            try
            {
                if (obj["StartMode"].ToString() == "Disabled")
                {
                    inParam = obj.GetMethodParameters("ChangeStartMode");
                    inParam["StartMode"] = "Manual";
                    outParam = obj.InvokeMethod("ChangeStartMode", inParam, null);
                }
                else if (obj["StartMode"].ToString() == "Manual")
                {
                    inParam = obj.GetMethodParameters("ChangeStartMode");
                    inParam["StartMode"] = "Disabled";
                    outParam = obj.InvokeMethod("ChangeStartMode", inParam, null);
                }
            }
            catch
            {
                throw;
            }
        }
        public static void startStopService(string hostname)
        {
            try
            {
                string serviceName = "RemoteRegistry";

                ServiceController serviceController = new ServiceController("Remote Registry", hostname);
                ConnectionOptions connectoptions = new ConnectionOptions();
                ManagementScope scope = new ManagementScope("\\\\" + hostname + "\\root\\CIMV2");
                scope.Options = connectoptions;
                
                SelectQuery query = new SelectQuery("select * from Win32_Service where name = '" + serviceName + "'");          //WMI query to be executed on the remote machine  
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query))
                {
                    ManagementObjectCollection collection = searcher.Get();
                    foreach (ManagementObject service in collection.OfType<ManagementObject>())
                    {
                        if (service["started"].Equals(true))
                        {
                            service.InvokeMethod("StopService", null);
                            serviceController.WaitForStatus(ServiceControllerStatus.Stopped);
                        }
                        else
                        {
                            service.InvokeMethod("StartService", null);
                            serviceController.WaitForStatus(ServiceControllerStatus.Running);
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
        }
        public static string[] pingHostname(string hostname)
        {
            Ping myPing = new Ping();
            string[] pingResults = new string[2] { "", "" };
            string[] failedResult = new string[1] { "TimeOut" };
            try
            {
                PingReply reply = myPing.Send(hostname, 10000);

                if (reply.Status.ToString() == "Success")
                {
                    pingResults[0] = reply.Address.ToString();
                    pingResults[1] = reply.RoundtripTime.ToString();
                    return pingResults;     //Online PC
                }
                else
                    return null;            //Offline PC
            }
            catch
            {
                return failedResult;        //Invalid PC Name
            }
        }
        public static string[,] copyfiles(string newPc, string oldPc, string username, string item)
        {
            
            if (item == "Edge Bookmarks")
            {
                string[,] result = new string[2,3];                     //Edge Bookmark Transfer requires 3 files to copy over
                string[,] edge = copyPath(newPc, oldPc, username, item);
                
                for (int i = 0; i<3; i++)
                {
                    if(!(File.Exists(edge[0, i])))
                    {
                        Array.Clear(edge, i, 1);
                        result[0, i] = (i+1).ToString();                //Transfer fail will return a string number in a string array to indicate which file is failed to transfer
                    }
                    else
                        File.Copy(edge[0, i], edge[1, i], true);        //Transfer success will return a null to a string array
                }
                return result;
            }
            else
            {
                string[,] result = new string[2, 2];                    //Chrome Bookmark transfer requires 2 files to copy over
                string[,] chrome = copyPath(newPc, oldPc, username, item);
                for (int i = 0; i < 2; i++)
                {
                    if (!(File.Exists(chrome[0, i])))
                    {
                        Array.Clear(chrome, i, 1);
                        result[0, i] = (i+1).ToString();                //Transfer fail will return a string number in a string array to indicate which file is failed to transfer
                    }
                    else
                        File.Copy(chrome[0, i], chrome[1, i], true);    //Transfer success will return a null to a string array
                }
                return result;
            }
        }
        public static bool copyDirectory(string newPc, string oldPc, string username, string item)
        {
            if (item == "Quick Access")
            {
                string[,] quickAccess = copyPath(newPc, oldPc, username, item);
                var files = new DirectoryInfo(quickAccess[1, 0]).GetFiles("*.*");

                foreach (FileInfo file in files)
                    file.CopyTo(quickAccess[0, 0] + file.Name, true);

                return true;
            }
            else if (item == "Outlook Signatures")
            {
                string[,] outlookSignatures = copyPath(newPc, oldPc, username, item);

                if (!Directory.Exists(outlookSignatures[0, 0]))            //Copy directory has to overwrite the exsiting folder
                    Directory.CreateDirectory(outlookSignatures[0, 0]);

                if (!Directory.Exists(outlookSignatures[1, 0]))
                    return false;
                else
                {
                    var files = new DirectoryInfo(outlookSignatures[1, 0]).GetFiles("*.*");
                    foreach (FileInfo file in files)
                        file.CopyTo(outlookSignatures[0, 0] + file.Name, true);

                    return true;
                }
            }
            else
                return false;
        }
        public static string[,] confirmPath(string newPc, string oldPc, string username)    //Only to confirm both path(New and Old PCs) are existed
        {                                                                                   //If a user have not signed on both PC, then Path won't be existed
            string[,] path = new string[2, 1]
            {
                { @"\\" + oldPc + @"\c$\users\" + username},
                { @"\\" + newPc + @"\c$\users\" + username}
            };
            return path;
        }
        public static string[,] copyPath(string newPc, string oldPc, string username, string item)
        {
            if (item == "Edge Bookmarks")
            {
                string[,] path = new string[2, 3]
                {
                    { @"\\" + oldPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks",
                        @"\\" + oldPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks.bak",
                        @"\\" + oldPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks.msbak" },
                    { @"\\" + newPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks",
                        @"\\" + newPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks.bak",
                        @"\\" + newPc + @"\c$\users\" + username + @"\appdata\local\microsoft\edge\user data\default\bookmarks.msbak" }
                };
                return path;
            }
            else if (item == "Chrome Bookmarks")
            {
                string[,] path = new string[2, 2]
                {
                    { @"\\" + oldPc + @"\c$\users\" + username + @"\appdata\local\google\chrome\user data\default\bookmarks",
                        @"\\" + oldPc + @"\c$\users\" + username + @"\appdata\local\google\chrome\user data\default\bookmarks.bak",},
                    { @"\\" + newPc + @"\c$\users\" + username + @"\appdata\local\google\chrome\user data\default\bookmarks",
                        @"\\" + newPc + @"\c$\users\" + username + @"\appdata\local\google\chrome\user data\default\bookmarks.bak",}
                };
                return path;
            }
            else if (item == "Quick Access")
            {
                string[,] path = new string[2, 1]
                {
                    { @"\\" + newPc + @"\c$\users\" + username + @"\appdata\roaming\microsoft\Windows\Recent\automaticDestinations\"},
                    {@"\\" + oldPc + @"\c$\users\" + username + @"\appdata\roaming\microsoft\Windows\Recent\automaticDestinations"}
                };
                return path;
            }
            else if (item == "Outlook Signatures")
            {
                string[,] path = new string[2, 1]
                {
                    { @"\\" + newPc + @"\c$\users\" + username + @"\appdata\roaming\microsoft\Signatures\"},
                    {@"\\" + oldPc + @"\c$\users\" + username + @"\appdata\roaming\microsoft\Signatures"}
                };
                return path;
            }
            else
                return null;
        }
    }
}
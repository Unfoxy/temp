using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
//using CTMS.ServerBaseline.SCCM.ClassLibrary;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;
using System.Management;
using System.IO;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Collections;
using System.Diagnostics.Contracts;
using System.Threading;

namespace consoleTesting
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DirectoryEntry ldapConnection = new DirectoryEntry("");
                ldapConnection.Path = "LDAP://OU=Users,OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
                //directoryEntry.Path = "LDAP://OU=Users,OU=BMC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
                //directoryEntry.Path = "LDAP://OU=Users,OU=CHO,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
                //directoryEntry.Path = "LDAP://OU=Users,OU=LAR,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
                //directoryEntry.Path = "LDAP://OU=Users,OU=HDC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
                ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

                ////Memo userAccontControl changes to 544 after enabling. It has to be 512
                //General
                DirectoryEntry childEntry = ldapConnection.Children.Add("CN=" + "Doe" + "\\" + ", " + "John" + " " + "T", "user");      //CN
                childEntry.Properties["givenName"].Value = "John";                                                                      //First name
                childEntry.Invoke("Put", new object[] { "Initials", "T" });                                                             //Initials
                childEntry.Properties["sn"].Value = "Doe";                                                                              //Last name
                childEntry.Properties["displayName"].Value = "Doe, John";                                                               //Display name
                childEntry.Invoke("Put", new object[] { "Description", "rAM - MVW - League City. Texas" });                             //Description
                childEntry.Properties["physicalDeliveryOfficeName"].Value = "MVW: ";                                                    //Office
                childEntry.Properties["telephoneNumber"].Value = "281-535-";                                                            //Telephone nuber
                childEntry.Properties["mail"].Value = "john.doe@ineos.com";                                                             //E-mail
                //Address
                childEntry.Properties["streetAddress"].Value = "2600 South Shore Blvd.";                                                //Street
                childEntry.Properties["l"].Value = "League City";                                                                       //City
                childEntry.Properties["st"].Value = "Texas";                                                                            //State/province
                childEntry.Properties["postalCode"].Value = "77573";                                                                    //Zip/postal code
                childEntry.Properties["c"].Value = "US";                                                                                //Country/region
                childEntry.Properties["co"].Value = "United States";                                                                    //Country/region
                //Account
                childEntry.Properties["userPrincipalName"].Value = "john.doe@ineos.com";                                                //User logon name
                childEntry.Properties["samAccountName"].Value = "jtd13153";                                                             //User logon name (pre-Windows 2000):
                //Profille
                childEntry.Properties["homeDirectory"].Value = "\\\\in1\\opu\\users\\jtd13153";                                         //Home folder
                childEntry.Properties["homeDrive"].Value = "H:";                                                                        //Connect
                //Organization
                childEntry.Properties["company"].Value = "INEOS O&P USA";                                                               //Company
                //Attributes
                childEntry.Properties["extensionAttribute1"].Value = "MVW";
                childEntry.Properties["extensionAttribute10"].Value = "EMPLOYEE";
                childEntry.Properties["extensionAttribute11"].Value = "GMAIL_SUB";
                childEntry.Properties["extensionAttribute12"].Value = "OPUSA_O365";
                childEntry.Properties["extensionAttribute13"].Value = "o365";                                                           //Create a prompt for ;OKTA
                childEntry.Properties["extensionAttribute2"].Value = "OP USA";
                childEntry.Properties["extensionAttribute9"].Value = "IN1_IN1_E3";
                childEntry.Properties["mailNickname"].Value = "john.doe";

                childEntry.CommitChanges();
                ldapConnection.CommitChanges();
                //childEntry.Invoke("SetPassword", new object[] { "Ineos2023" });
                //childEntry.CommitChanges();
                Console.WriteLine("Creation Success, ");
                Console.WriteLine("Press any key to copy membership");
                Console.ReadLine();

                //---------------------------------------------------------Duplicate Membership-------------------------------------------------------

                PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "in1.ad.innovene.com", "OU=Users,OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com");

                UserPrincipal sourceUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, "yxl13153");
                UserPrincipal destinationUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, "jtd13153");

                if (sourceUser != null && destinationUser != null)
                {
                    var sourceGroups = sourceUser.GetGroups();

                    var destinationGroups = destinationUser.GetGroups();

                    foreach (Principal sourceGroup in sourceGroups)
                    {
                        if (!destinationGroups.Contains(sourceGroup))
                        {
                            GroupPrincipal destinationGroup = sourceGroup as GroupPrincipal;
                            destinationGroup.Members.Add(destinationUser);
                            destinationGroup.Save();
                        }
                    }
                }

                Console.WriteLine();
                Console.WriteLine("Success");
                Console.ReadLine();
            }
            catch
            {
                Console.WriteLine();
                Console.WriteLine("Failed");
                Console.ReadLine();
            }
        }
    }
}
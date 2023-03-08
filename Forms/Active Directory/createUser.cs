using System;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.DirectoryServices.AccountManagement;
using System.CodeDom.Compiler;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Runtime.CompilerServices;

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class createUser : Form
    {
        public createUser()
        {
            InitializeComponent();
            txtCreateUserContractorCompany.BackColor = Color.FromArgb(220, 220, 220);
            txtCreateUserContractorEmail.BackColor = Color.FromArgb(220, 220, 220);
            disposeMembershipsCount = 0;
            disposeMembershipsGroup = new string[1000];
        }
        private void comboBoxCreateUserUserType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxCreateUserUserType.Text == "Contractor")
            {
                txtCreateUserContractorCompany.Enabled = true;
                txtCreateUserContractorCompany.BackColor = Color.FromArgb(32, 33, 36);
                txtCreateUserContractorEmail.Enabled = true;
                txtCreateUserContractorEmail.BackColor = Color.FromArgb(32, 33, 36);
            }
            else if(comboBoxCreateUserUserType.Text == "External")
            {
                txtCreateUserContractorCompany.Enabled = true;
                txtCreateUserContractorCompany.BackColor = Color.FromArgb(32, 33, 36);
                txtCreateUserContractorEmail.Enabled = false;
                txtCreateUserContractorEmail.BackColor = Color.FromArgb(220, 220, 220);
            }
            else if (comboBoxCreateUserUserType.Text == "Employee")
            {
                txtCreateUserContractorCompany.Enabled = false;
                txtCreateUserContractorCompany.BackColor = Color.FromArgb(220, 220, 220);
                txtCreateUserContractorEmail.Enabled = false;
                txtCreateUserContractorEmail.BackColor = Color.FromArgb(220, 220, 220);
            }
        }
        private void checkBoxO365MailEnabled_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateUserO365MailEnabled.CheckState == CheckState.Unchecked)
            {
                comboBoxCreateUserO365License.Enabled = false;
                comboBoxCreateUserO365License.BackColor = Color.FromArgb(220, 220, 220);
                if (comboBoxCreateUserUserType.Text == "Contractor")
                {
                    txtCreateUserContractorEmail.Enabled = true;
                    txtCreateUserContractorEmail.BackColor = Color.FromArgb(32, 33, 36);
                }
            }
            else
            {
                comboBoxCreateUserO365License.Enabled = true;
                comboBoxCreateUserO365License.BackColor = Color.FromArgb(32, 33, 36);
                if (comboBoxCreateUserUserType.Text == "Contractor")
                {
                    txtCreateUserContractorEmail.Enabled = false;
                    txtCreateUserContractorEmail.BackColor = Color.FromArgb(220, 220, 220);
                }
            }
        }
        string firstname { get; set; }
        string firstInitial { get; set; }
        string lastname { get; set; }
        string lastInitial { get; set; }
        string middleName { get; set; }
        string middleInitial { get; set; }
        string email { get; set; }
        string ntid { get; set; }
        string contractorCompanyName { get; set; }
        string officeNumber { get; set; }
        string telephoneNumber { get; set; }
        string jobTitle { get; set; }
        string department { get; set; }
        string managerNTID { get; set; }
        string extensionAttribute1 { get; set; }
        string extensionAttribute2 { get; set; }
        string extensionAttribute9 { get; set; }
        string extensionAttribute10 { get; set; }
        string extensionAttribute11 { get; set; }
        string extensionAttribute12 { get; set; }
        string extensionAttribute13 { get; set; }
        
        
        private void btnCreateUserGenerateNTID_Click(object sender, EventArgs e)
        {
            //Firstname
            firstname = txtCreateUserFirstname.Text;
            firstname = firstname.ToLower();
            firstname = char.ToUpper(firstname[0]) + firstname.Substring(1);    //Make first letter uppercase
            firstInitial = firstname.Substring(0, 1);
            //Lastname
            lastname = txtCreateUserLastname.Text;
            lastname = lastname.ToLower();
            lastname = char.ToUpper(lastname[0]) + lastname.Substring(1);       //Make first letter uppercase
            lastInitial = lastname.Substring(0, 1);
            //Middle Initial
            if (txtCreatUserMiddleInitial.Text == "")
                middleInitial = "X";
            else
            {
                middleName = txtCreatUserMiddleInitial.Text;
                middleName = middleName.ToUpper();
                middleInitial = middleName.Substring(0, 1);
            }
            //NTID
            Random rnd = new Random();
            int num = rnd.Next(100000);
            string ntidNum = num.ToString();
            ntid = firstInitial + middleInitial + lastInitial + ntidNum;
            ntid = ntid.ToLower();
            txtCreateUserNewUserAccount.Text = ntid;
            //Email
            if (comboBoxCreateUserUserType.Text == "Employee")
                email = firstname.ToLower() + "." + lastname.ToLower() + "@ineos.com";
            else if (comboBoxCreateUserUserType.Text == "External")
                email = firstname.ToLower() + "." + lastname.ToLower() + "@external.ineos.com";
            else if (comboBoxCreateUserUserType.Text == "Contractor" && checkBoxCreateUserO365MailEnabled.Checked)
                email = firstname.ToLower() + "." + lastname.ToLower() + "@external.ineos.com";
            else if (comboBoxCreateUserUserType.Text == "Contractor" && !(checkBoxCreateUserO365MailEnabled.Checked))
                email = txtCreateUserContractorEmail.Text;
            else
                email = "";
            //Contractor Company Name
            contractorCompanyName = txtCreateUserContractorCompany.Text;
            //Office Number
            officeNumber = comboBoxCreateUserSiteCode.Text + " " + txtCreateUserOffice.Text;
            //Telephone Number
            telephoneNumber = txtCreateUserPhone.Text;
            //Job Title
            jobTitle = txtlbCreateUserPosition.Text;
            //Department
            department = txtCreateUserDepartment.Text;
            //Manager
            managerNTID = txtCreateUserSupervisor.Text;

            //EA 1, 2, 9, 10, 11, 12, 13 Need to be assigned
            //EA 1, 9, 10 are combobox inputs
            if (comboBoxCreateUserSiteCode != null)                                             //EA 1
                extensionAttribute1 = comboBoxCreateUserSiteCode.Text;
            else
                extensionAttribute1 = "";
            if (checkBoxCreateUserO365MailEnabled.Checked)                                      //EA 9
                extensionAttribute9 = "IN1_IN1_" + comboBoxCreateUserO365License.Text;
            else
                extensionAttribute9 = "";
            if (comboBoxCreateUserUserType.Text != null)                                        //EA 10
                extensionAttribute10 = comboBoxCreateUserUserType.Text;
            else
                extensionAttribute10 = "";
            //EA 2, 11, 12 are manaul inputs
            extensionAttribute2 = "OP USA";                                                     //EA 2
            extensionAttribute11 = "GMAIL_SUB";                                                 //EA 11 *Exceptional IN1_SUB(MVW Treasury)
            extensionAttribute12 = "OPUSA_O365";                                                //EA 12 *Exceptional OPUSA_G_O365(Mailbox Account), RAM_O365(MVW Treasury)
            //EA 13 is checkbox values
            if(checkBoxCreateUserO365MailEnabled.Checked)
            {
                extensionAttribute13 = "o365";
            }
            if(checkBoxCreateUserEnableOkta.Checked)
            {
                if(checkBoxCreateUserO365MailEnabled.Checked)
                    extensionAttribute13 = "o365;OKTA";
                else
                    extensionAttribute13 = "OKTA";
            }
        }
        private void btnCreateUserCreateAccount_Click(object sender, EventArgs e)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry("");
            if (extensionAttribute1 == "MVW")
                ldapConnection.Path = "LDAP://OU=Users,OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (extensionAttribute1 == "BMC")
                ldapConnection.Path = "LDAP://OU=BMC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (extensionAttribute1 == "CHO")
                ldapConnection.Path = "LDAP://OU=CHO,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (extensionAttribute1 == "LAR")
                ldapConnection.Path = "LDAP://OU=LAR,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (extensionAttribute1 == "HDC")
                ldapConnection.Path = "LDAP://OU=HDC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            //Creation Start
            DirectoryEntry childEntry = ldapConnection.Children.Add("CN=" + lastname + "\\" + ", " + firstname + " " + middleInitial, "user");
            if (extensionAttribute10 == "Employee")                                                                                    //Display name
                childEntry.Properties["displayName"].Value = lastname + ", " + firstname + " " + middleInitial;
            else if (extensionAttribute10 == "External" || extensionAttribute10 == "Contractor")
                childEntry.Properties["displayName"].Value = lastname + ", " + firstname + " " + middleInitial + " (" + contractorCompanyName + ")";
            if (middleInitial == "X")                                                                                                  //No Middle Initial
            {
                childEntry = ldapConnection.Children.Add("CN=" + lastname + "\\" + ", " + firstname, "user");
                if (extensionAttribute10 == "Employee")                                                                                
                    childEntry.Properties["displayName"].Value = lastname + ", " + firstname;
                else if (extensionAttribute10 == "External" || extensionAttribute10 == "Contractor")
                    childEntry.Properties["displayName"].Value = lastname + ", " + firstname + " (" + contractorCompanyName + ")";
            }
            childEntry.Properties["givenName"].Value = firstname;                                                                      //First name
            childEntry.Invoke("Put", new object[] { "Initials", middleInitial });                                                      //Initials
            childEntry.Properties["sn"].Value = lastname;                                                                              //Last name
            

            if (extensionAttribute1 == "MVW")                                                                                          //Description
                childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - League City. Texas" });
            else if (extensionAttribute1 == "BMC")
                childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - LaPorte. Texas" });
            else if (extensionAttribute1 == "CHO")
                childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - Alvin. Texas" });
            else if (extensionAttribute1 == "LAR")
                childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - Carson. Texas" });
            else if (extensionAttribute1 == "HDC")
                childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - Houston. Texas" });
            childEntry.Properties["physicalDeliveryOfficeName"].Value = officeNumber;                                                  //Office
            childEntry.Properties["telephoneNumber"].Value = telephoneNumber;                                                          //Telephone number
            childEntry.Properties["mail"].Value = email;                                                                               //E-mail
            //Address
            if (extensionAttribute1 == "MVW")
            {                                                                                                                          //MVW Address
                childEntry.Properties["streetAddress"].Value = "2600 South Shore Blvd";                                                //Street
                childEntry.Properties["l"].Value = "League City";                                                                      //City
                childEntry.Properties["st"].Value = "Texas";                                                                           //State/province
                childEntry.Properties["postalCode"].Value = "77573";                                                                   //Zip/postal code
                childEntry.Properties["c"].Value = "US";                                                                               //Country/region
                childEntry.Properties["co"].Value = "United States";                                                                   //Country/region
            }
            else if (extensionAttribute1 == "BMC")                                                                                     //BMC Address
            {
                childEntry.Properties["streetAddress"].Value = "1230 Independence Pkwy";                                               
                childEntry.Properties["l"].Value = "La Porte";                                                                         
                childEntry.Properties["st"].Value = "Texas";                                                                           
                childEntry.Properties["postalCode"].Value = "77571";                                                                   
                childEntry.Properties["c"].Value = "US";                                                                               
                childEntry.Properties["co"].Value = "United States";                                                                   
            }
            else if (extensionAttribute1 == "CHO")                                                                                     //CHO Address
            {
                childEntry.Properties["streetAddress"].Value = "2 Miles SO of FM2917 on FM2004";                                       
                childEntry.Properties["l"].Value = "Alvin";                                                                            
                childEntry.Properties["st"].Value = "Texas";                                                                           
                childEntry.Properties["postalCode"].Value = "77511";                                                                   
                childEntry.Properties["c"].Value = "US";                                                                               
                childEntry.Properties["co"].Value = "United States";                                                                   
            }
            else if (extensionAttribute1 == "LAR")                                                                                     //LAR Address
            {
                childEntry.Properties["streetAddress"].Value = "2384 East 223rd St";
                childEntry.Properties["l"].Value = "Carson";
                childEntry.Properties["st"].Value = "CA";
                childEntry.Properties["postalCode"].Value = "90745";
                childEntry.Properties["c"].Value = "US";
                childEntry.Properties["co"].Value = "United States";
            }
            else if (extensionAttribute1 == "HDC")                                                                                     //HDC Address
            {
                childEntry.Properties["st"].Value = "Texas";
                childEntry.Properties["c"].Value = "US";
                childEntry.Properties["co"].Value = "United States";
            }
            //Account                                                                                                               
            childEntry.Properties["userPrincipalName"].Value = email;                                                                  //User logon name
            childEntry.Properties["samAccountName"].Value = ntid;                                                                      //User logon name (pre-Windows 2000):
            //Profille                                                                                                              
            childEntry.Properties["homeDirectory"].Value = "\\\\in1\\opu\\users\\" + ntid;                                             //Home folder
            childEntry.Properties["homeDrive"].Value = "H:";                                                                           //Connect
            //Organization                                                                                                             
            childEntry.Properties["title"].Value = jobTitle;                                                                           //Job title
            childEntry.Properties["department"].Value = department;                                                                    //Department
            childEntry.Properties["company"].Value = "INEOS O&P USA";                                                                  //Company
            if (txtCreateUserSupervisor.Text != "")                                                                                    //Manager
            {
                string[] manager = Functions.GetAD(txtCreateUserSupervisor.Text);
                string displayName = manager[0].Replace(",", "\\,");
                string site = manager[3];
                childEntry.Properties["manager"].Value = "CN=" + displayName + ",OU=Users,OU=" + site + ",OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            }
            //Attributes
            childEntry.Properties["extensionAttribute1"].Value = extensionAttribute1;
            childEntry.Properties["extensionAttribute2"].Value = extensionAttribute2;
            childEntry.Properties["extensionAttribute9"].Value = extensionAttribute9;  
            childEntry.Properties["extensionAttribute10"].Value = extensionAttribute10;
            childEntry.Properties["extensionAttribute11"].Value = extensionAttribute11;
            childEntry.Properties["extensionAttribute12"].Value = extensionAttribute12;
            childEntry.Properties["extensionAttribute13"].Value = extensionAttribute13;
            childEntry.Properties["mailNickname"].Value = firstname.ToLower() + "." + lastname.ToLower();

            childEntry.CommitChanges();
            ldapConnection.CommitChanges();
            //childEntry.Invoke("SetPassword", new object[] { "Ineos2023" });
            //childEntry.CommitChanges();
            var (adminMemberships, adminMembershipsCount) = Functions.copyMembership(disposeMembershipsGroup, disposeMembershipsCount, txtCreateUserCopyMembership.Text, txtCreateUserNewUserAccount.Text);
            MessageBox.Show("Membership Copy Fail "
                                                             + "\n" + adminMemberships[0],
                                   "Membership Copy Fail", MessageBoxButtons.OK, MessageBoxIcon.Warning).ToString();
        }
        string disposeMembership { get; set; }
        string[] disposeMembershipsGroup { get; set; }
        int disposeMembershipsCount { get; set; }
        string[] copyMemberships { get; set; }
        int copyMembershipsCount { get; set; }
        private void btnCreateUserMembershipsDisplay_Click(object sender, EventArgs e)
        {
            string ntid = txtCreateUserCopyMembership.Text;//"yxl13153";
            (copyMemberships, copyMembershipsCount) = Functions.displayMembership(ntid);

            rtxtCreateUserMemberships.AppendText("Current Memberships");
            for (int i = 0; i < copyMembershipsCount; i++)
            {
                if (i >= 9)
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i + 1) + ". " + copyMemberships[i]);
                else
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + "0" + (i + 1) + ". " + copyMemberships[i]);
            }
        }
        public void btnCreateUserMembershipsRemove_Click(object sender, EventArgs e)
        {
            (copyMemberships, copyMembershipsCount) = Functions.removeMembership(copyMemberships, copyMembershipsCount, disposeMembership);

            (disposeMembershipsGroup, disposeMembershipsCount) = Functions.disposeMemberships(disposeMembershipsGroup, disposeMembershipsCount, disposeMembership);

            rtxtCreateUserMemberships.Clear();
            rtxtCreateUserMemberships.Focus();

            rtxtCreateUserMemberships.SelectionColor = Color.FromArgb(255, 255, 128);
            rtxtCreateUserMemberships.AppendText("Removed Memberships");
            for (int i = 0; i < disposeMembershipsCount; i++)
            {
                if (i >= 9)
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i + 1) + ". " + disposeMembershipsGroup[i]);
                else
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + "0" + (i + 1) + ". " + disposeMembershipsGroup[i]);
            }

            rtxtCreateUserMemberships.AppendText(Environment.NewLine);
            rtxtCreateUserMemberships.AppendText(Environment.NewLine);

            rtxtCreateUserMemberships.AppendText("Copied Memberships");
            for (int i = 0; i < copyMembershipsCount; i++)
            {
                if (i >= 9)
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i + 1) + ". " + copyMemberships[i]);
                else
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + "0" + (i + 1) + ". " + copyMemberships[i]);
            }

        }
        public void rtxtCreateUserMemberships_MouseDown(object sender, MouseEventArgs e)
        {
            int index = rtxtCreateUserMemberships.GetCharIndexFromPosition(e.Location);
            int line = rtxtCreateUserMemberships.GetLineFromCharIndex(index);

            int lineStartIndex = rtxtCreateUserMemberships.GetFirstCharIndexFromLine(line);
            int lineLength = rtxtCreateUserMemberships.Lines[line].Length;
            rtxtCreateUserMemberships.Select(lineStartIndex, lineLength);

            string selectedLine = rtxtCreateUserMemberships.SelectedText;
            selectedLine = selectedLine.Trim();
            disposeMembership = selectedLine.Remove(0, 4);
        }
    }
}
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

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class createUser : Form
    {
        public createUser()
        {
            InitializeComponent();
            txtCreateUserContractor.BackColor = Color.FromArgb(220, 220, 220);
            txtCreateUserContractor.BackColor = Color.FromArgb(220, 220, 220);
            disposeMembershipsCount = 0;
            disposeMembershipsGroup = new string[1000];
        }
        private void checkBoxEmployee_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateUserEmployee.Checked)
            {
                checkBoxCreateUserContractor.Checked = false;
                checkBoxCreateUserExternal.Checked = false;
            }
        }

        private void checkBoxExternal_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateUserExternal.Checked)
            {
                checkBoxCreateUserEmployee.Checked = false;
                checkBoxCreateUserContractor.Checked = false;
            }
        }
        private void checkBoxContractor_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateUserContractor.Checked)
            {
                checkBoxCreateUserEmployee.Checked = false;
                checkBoxCreateUserExternal.Checked = false;
            }
            if (checkBoxCreateUserContractor.CheckState == CheckState.Unchecked)
            {
                txtCreateUserContractor.Enabled = false;
                txtCreateUserContractor.BackColor = Color.FromArgb(220, 220, 220);
            }
            else
            {
                txtCreateUserContractor.Enabled = true;
                txtCreateUserContractor.BackColor = Color.FromArgb(32, 33, 36);
            }
        }

        private void checkBoxO365MailEnabled_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxCreateUserO365MailEnabled.CheckState == CheckState.Unchecked)
            {
                comboBoxCreateUserO365License.Enabled = false;
                comboBoxCreateUserO365License.BackColor = Color.FromArgb(220, 220, 220);
            }
            else
            {
                comboBoxCreateUserO365License.Enabled = true;
                comboBoxCreateUserO365License.BackColor = Color.FromArgb(32, 33, 36);
            }
        }
        string firstname { get; set; }
        string firstInitial { get; set; }
        string lastname { get; set; }
        string lastInitial { get; set; }
        string middleName { get; set; }
        string middleInitial { get; set; }
        string email { get; set; }
        string extensionAttribute1 { get; set; }
        string extensionAttribute9 { get; set; }
        string ntid { get; set; }
        private void btnCreateUserGenerateNTID_Click(object sender, EventArgs e)
        {
            firstname = txtCreateUserFirstname.Text;
            firstname = firstname.ToLower();
            firstname = char.ToUpper(firstname[0]) + firstname.Substring(1);
            firstInitial = firstname.Substring(0, 1);

            lastname = txtCreateUserLastname.Text;
            lastname = lastname.ToLower();
            lastname = char.ToUpper(lastname[0]) + lastname.Substring(1);
            lastInitial = lastname.Substring(0, 1);

            middleName = txtCreatUserMiddleInitial.Text;
            middleName = middleName.ToUpper();
            if (middleName == null)
                middleName = "X";
            middleInitial = middleName.Substring(0, 1);

            Random rnd = new Random();
            int num = rnd.Next(100000);
            string ntidNum = num.ToString();
            ntid = firstInitial + middleInitial + lastInitial + ntidNum;
            ntid = ntid.ToLower();

            if (checkBoxCreateUserEmployee.Checked)                                             //Email
                email = firstname.ToLower() + "." + lastname.ToLower() + "@ineos.com";
            else if (checkBoxCreateUserExternal.Checked)
                email = firstname.ToLower() + "." + lastname.ToLower() + "@external.ineos.com";
            else if (checkBoxCreateUserContractor.Checked)
                email = txtCreateUserContractor.Text;
            else
                email = "";

            if (checkBoxCreateUserO365MailEnabled.Checked)                                      //EA9
                extensionAttribute9 = "IN1_IN1_" + comboBoxCreateUserO365License.Text;

            if (comboBoxCreateUserSiteCode != null)                                             //EA1
                extensionAttribute1 = comboBoxCreateUserSiteCode.Text;
            else
                extensionAttribute1 = "";

            txtCreateUserNewUserAccount.Text = ntid;
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
            for (int i=0; i < copyMembershipsCount; i++)
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
            for (int i=0; i < disposeMembershipsCount; i++)
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

        private void btnCreateUserCreateAccount_Click(object sender, EventArgs e)
        {
            DirectoryEntry ldapConnection = new DirectoryEntry("");
            if (comboBoxCreateUserSiteCode.Text == "MVW")
                ldapConnection.Path = "LDAP://OU=Users,OU=MVW,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (comboBoxCreateUserSiteCode.Text == "BMC")
                ldapConnection.Path = "LDAP://OU=BMC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (comboBoxCreateUserSiteCode.Text == "CHO")
                ldapConnection.Path = "LDAP://OU=CHO,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (comboBoxCreateUserSiteCode.Text == "LAR")
                ldapConnection.Path = "LDAP://OU=LAR,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            else if (comboBoxCreateUserSiteCode.Text == "HDC")
                ldapConnection.Path = "LDAP://OU=HDC,OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            //Creation Start
            DirectoryEntry childEntry = ldapConnection.Children.Add("CN=" + lastname + "\\" + ", " + firstname + " " + middleInitial, "user");
            childEntry.Properties["givenName"].Value = firstname;                                                                      //First name
            childEntry.Invoke("Put", new object[] { "Initials", middleInitial });                                                             //Initials
            childEntry.Properties["sn"].Value = lastname;                                                                              //Last name
            childEntry.Properties["displayName"].Value = lastname + "\\" + ", " + firstname;                                                                   //Display name
            childEntry.Invoke("Put", new object[] { "Description", "rAM - " + extensionAttribute1 + " - League City. Texas" });                             //Description
            childEntry.Properties["physicalDeliveryOfficeName"].Value = extensionAttribute1 + ": ";                                                    //Office
            childEntry.Properties["telephoneNumber"].Value = "281-535-";                                                            //Telephone nuber
            childEntry.Properties["mail"].Value = email;                                                                            //E-mail
            //Address                                                                                                             //Address
            childEntry.Properties["streetAddress"].Value = "2600 South Shore Blvd.";                                                //Street
            childEntry.Properties["l"].Value = "League City";                                                                       //City
            childEntry.Properties["st"].Value = "Texas";                                                                            //State/province
            childEntry.Properties["postalCode"].Value = "77573";                                                                    //Zip/postal code
            childEntry.Properties["c"].Value = "US";                                                                                //Country/region
            childEntry.Properties["co"].Value = "United States";                                                                    //Country/region
            //Account                                                                                                             //Account
            childEntry.Properties["userPrincipalName"].Value = email;                                                //User logon name
            childEntry.Properties["samAccountName"].Value = ntid;                                                             //User logon name (pre-Windows 2000):
            //Profille                                                                                                              //Profille
            childEntry.Properties["homeDirectory"].Value = "\\\\in1\\opu\\users\\" + ntid;                                         //Home folder
            childEntry.Properties["homeDrive"].Value = "H:";                                                                        //Connect
            //Organization                                                                                                             //Organization
            childEntry.Properties["company"].Value = "INEOS O&P USA";                                                               //Company
                                                                                                                                    //Attributes
            childEntry.Properties["extensionAttribute1"].Value = extensionAttribute1;
            childEntry.Properties["extensionAttribute10"].Value = "EMPLOYEE";
            childEntry.Properties["extensionAttribute11"].Value = "GMAIL_SUB";
            childEntry.Properties["extensionAttribute12"].Value = "OPUSA_O365";
            childEntry.Properties["extensionAttribute13"].Value = "o365";                                                           //Create a prompt for ;OKTA
            childEntry.Properties["extensionAttribute2"].Value = "OP USA";
            childEntry.Properties["extensionAttribute9"].Value = "IN1_IN1_E3";
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
    }
}




//// Assuming you have a RichTextBox control named richTextBox1
//List<string> lines = new List<string>();
//string[] outputLine = new string[1000];
//int count = 0;
//foreach (string line in rtxtCreateUserMemberships.Lines)
//{
//    lines.Add(line);
//    outputLine[count] = lines[count];
//    outputLine[count].Remove(0, 3);
//    count++;
//}

//// Now the lines of text are stored in the 'lines' list
//// You can access each line using the index of the list, e.g.:
////string firstLine = lines[0];
////string secondLine = lines[1];
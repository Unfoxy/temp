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

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class createUser : Form
    {
        public createUser()
        {
            InitializeComponent();
            txtCreateUserContractor.BackColor = Color.FromArgb(220, 220, 220);
            txtCreateUserContractor.BackColor = Color.FromArgb(220, 220, 220);
            disposeMembershipCount = 0;
            disposeTesting = new string[1000];
        }

        private void checkBoxContractor_CheckedChanged(object sender, EventArgs e)
        {
            //checkBoxCreateUserExternal.Checked = !checkBoxCreateUserContractor.Checked;
            //checkBoxCreateUserEmployee.Checked = !checkBoxCreateUserContractor.Checked;

            //checkBoxCreateUserContractor.Checked = !checkBoxCreateUserContractor.Checked;
            //checkBoxCreateUserContractor.Checked = !checkBoxCreateUserExternal.Checked;

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

        private void checkBoxEmployee_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxCreateUserExternal.Checked = !checkBoxCreateUserEmployee.Checked;
            checkBoxCreateUserContractor.Checked = !checkBoxCreateUserEmployee.Checked;
        }

        private void checkBoxExternal_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxCreateUserEmployee.Checked = !checkBoxCreateUserExternal.Checked;
            checkBoxCreateUserContractor.Checked = !checkBoxCreateUserExternal.Checked;
        }

        private void btnCreateUserGenerateNTID_Click(object sender, EventArgs e)
        {
            string firstname = txtCreateUserFirstname.Text;
            string lastname = txtCreateUserLastname.Text;
            string middleName = txtCreatUserMiddleInitial.Text;
            string email;
            string extensionAttribute9;
            string extensionAttribute1;
            string ntid;

            if (middleName == null)
                middleName = "x";                                    

            Random rnd = new Random();
            int num = rnd.Next(100000);
            string ntidNum = num.ToString();

            string firstInitial = firstname.Substring(0, 1);
            string middleInitial = middleName.Substring(0, 1);
            string lastInitial = lastname.Substring(0, 1);
            ntid = firstInitial + middleInitial + lastInitial + ntidNum.ToLower();

            if (checkBoxCreateUserEmployee.Checked)                                             //Email
                email = firstname + lastname + "@ineos.com";
            else if (checkBoxCreateUserExternal.Checked)
                email = firstname + lastname + "@external.ineos.com";
            else if (checkBoxCreateUserContractor.Checked)
                email = txtCreateUserContractor.Text;

            if (checkBoxCreateUserO365MailEnabled.Checked)                                      //EA9
                extensionAttribute9 = "IN1_IN1_" + comboBoxCreateUserO365License.Text;

            if (comboBoxCreateUserSiteCode != null)                                             //EA1
                extensionAttribute1 = comboBoxCreateUserSiteCode.Text;

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








        }
        string[] disposeTesting { get; set; }
        string disposeMembership { get; set; }
        int disposeMembershipCount { get; set; }
        string[] displayMemberships { get; set; }
        int membershipsCount { get; set; }
        private void btnCreateUserMembershipsDisplay_Click(object sender, EventArgs e)
        {
            string ntid = "yxl13153";
            (displayMemberships, membershipsCount) = Functions.displayMembership(null, ntid);

            for (int i=0; i < membershipsCount; i++)
            {
                if (i >= 9)
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i + 1) + ". " + displayMemberships[i]);
                else
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + "0" + (i + 1) + ". " + displayMemberships[i]);
            }
        }
        public void btnCreateUserMembershipsRemove_Click(object sender, EventArgs e)
        {
            //string ntid = "yxl13153";
            //var (displayMemberships, membershipsCount) = Functions.displayMembership(null, ntid);
            (displayMemberships, membershipsCount) = Functions.removeMembership(displayMemberships, membershipsCount, disposeMembership);
            
            (disposeTesting, disposeMembershipCount) = Functions.disposeMemberships(disposeTesting, disposeMembershipCount, disposeMembership);

            //testing[disposeMembershipCount] = disposeMembership;
            //disposeMembershipCount++;


            rtxtCreateUserMemberships.Clear();
            rtxtCreateUserMemberships.Focus();


            //rtxtCreateUserMemberships.SelectionColor = Color.FromArgb(255, 255, 128) ;
            //rtxtCreateUserMemberships.AppendText(disposeMembership + " is removed.");
            //rtxtCreateUserMemberships.AppendText(Environment.NewLine);

            rtxtCreateUserMemberships.SelectionColor = Color.FromArgb(255, 255, 128);
            rtxtCreateUserMemberships.AppendText("Removed Memberships");
            for (int i=0; i < disposeMembershipCount; i++)
                rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i+1) + ". " + disposeTesting[i]);
            
            rtxtCreateUserMemberships.AppendText(Environment.NewLine);

            for (int i = 0; i < membershipsCount; i++)
            {
                if (i >= 9)
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + (i + 1) + ". " + displayMemberships[i]);
                else
                    rtxtCreateUserMemberships.AppendText(Environment.NewLine + "0" + (i + 1) + ". " + displayMemberships[i]);
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
////display membership
//PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "in1.ad.innovene.com", "OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com");
//UserPrincipal sourceUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, "yxl13153");


//if (sourceUser != null)
//{
//    var sourceGroups = sourceUser.GetGroups();

//    int count = 0;
//    rtxtCreateUserMemberships.AppendText("Memberships");
//    foreach (Principal sourceGroup in sourceGroups)
//    {
//        rtxtCreateUserMemberships.AppendText(Environment.NewLine + (count + 1) + ". " + sourceGroup.Name);
//        count++;
//        string temp;
//        temp = sourceGroup.Name;

//    }
//}


//// Copy Membership
//PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, "in1.ad.innovene.com", "OU=rAM,OU=Client,DC=in1,DC=ad,DC=innovene,DC=com");
//UserPrincipal sourceUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, "yxl13153");
////UserPrincipal destinationUser = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, txtCreateUserNewUserAccount.Text);

//if (sourceUser != null)//&& destinationUser != null
//{
//    var sourceGroups = sourceUser.GetGroups();
//    //var destinationGroups = destinationUser.GetGroups();
//    int count = 0;
//    rtxtCreateUserMemberships.AppendText("Memberships");
//    foreach (Principal sourceGroup in sourceGroups)
//    {
//        rtxtCreateUserMemberships.AppendText(Environment.NewLine + (count + 1) + ". " + sourceGroup.Name);
//        count++;
//        string temp;
//        temp = sourceGroup.Name;
//        //if (!destinationGroups.Contains(sourceGroup))
//        //{

//        //    GroupPrincipal destinationGroup = sourceGroup as GroupPrincipal;
//        //    destinationGroup.Members.Add(destinationUser);
//        //    if (destinationGroup.DistinguishedName == "CN=" + destinationGroup.Name + "," + "OU=RG,OU=rAM,OU=Admin,DC=in1,DC=ad,DC=innovene,DC=com")
//        //    {
//        //        destinationGroup.Dispose();
//        //    }
//        //    else
//        //        destinationGroup.Save();
//        //}
//    }
//}



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
using desktopDashboard___Y_Lee;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class dataMigration : Form
    {
        public dataMigration()
        {
            InitializeComponent();
        }
        private void btnDataMigrationOK_Click(object sender, EventArgs e)
        {
            string newPc = txtDataMigrationNewPc.Text;
            string oldPc = txtDataMigrationOldPc.Text;
            string username = txtDataMigrationUserId.Text;
            string item = comboxDataMigration.Text.ToString();

            try
            {
                string[] usernameAD = Functions.GetAD(username);
                if (usernameAD == null)
                    rtxtDataMigration.Text = "Invalid Username Entry\n" + "\nUsername: '" + username.ToUpper() + "' is Incorrect";
                else
                {
                    if (item == "")
                    {
                        rtxtDataMigration.Text = "Please Select Object\n"
                                                + "\n1. Edge Bookmarks"
                                                + "\n2. Chrome Bookmarks"
                                                + "\n3. Outlook Signatures"
                                                + "\n4. Quick Access";
                    }
                    else
                    {
                        string answer = MessageBox.Show("Please Confirm Again "
                                                             + "\nNew PC: " + newPc.ToUpper()
                                                             + "\nOld PC: " + oldPc.ToUpper()
                                                             + "\nUser ID: " + username.ToUpper()
                                                             + "\nUser name: " + usernameAD[0],
                                   "Desktop Dashboard", MessageBoxButtons.YesNo, MessageBoxIcon.Warning).ToString();

                        if (answer == "Yes")
                        {
                            string[] newPcPing = Functions.pingHostname(newPc);
                            string[] oldPcPing = Functions.pingHostname(oldPc);
                            string[,] pathConfirm = Functions.confirmPath(newPc, oldPc, username);

                            if (Directory.Exists(pathConfirm[0, 0]) || Directory.Exists(pathConfirm[1, 0]))         // path doesn't exist if user doesn't sign on at least 1 time
                            {
                                if (newPcPing != null && oldPcPing != null)
                                {
                                    if (item == "Chrome Bookmarks")
                                    {
                                        string[,] result = Functions.copyfiles(newPc, oldPc, username, item);
                                        string[] reference = new string[2] { "Bookmarks", "Bookmarks.bak" };

                                        rtxtDataMigration.Text = "Transferring 2 files - Chrome";
                                        rtxtDataMigration.AppendText(Environment.NewLine + "1. " + reference[0]);
                                        rtxtDataMigration.AppendText(Environment.NewLine + "2. " + reference[1]);
                                        rtxtDataMigration.AppendText(Environment.NewLine);
                                        rtxtDataMigration.AppendText(Environment.NewLine);

                                        for (int i = 0; i < 2; i++)
                                        {
                                            if (result[0, i] == (i + 1).ToString())
                                                rtxtDataMigration.AppendText(Environment.NewLine + (i + 1) + ". " + reference[i] + " - is Missing from '" + oldPc + "'");
                                            else
                                                rtxtDataMigration.AppendText(Environment.NewLine + (i + 1) + ". " + reference[i] + " - Transfer Completed.");
                                        }
                                    }
                                    else if (item == "Edge Bookmarks")
                                    {
                                        string[,] result = Functions.copyfiles(newPc, oldPc, username, item);
                                        string[] reference = new string[3] { "Bookmarks", "Bookmarks.bak", "Bookmarks.msbak" };

                                        rtxtDataMigration.Text = "Transferring 3 files - Edge";
                                        rtxtDataMigration.AppendText(Environment.NewLine + "1. " + reference[0]);
                                        rtxtDataMigration.AppendText(Environment.NewLine + "2. " + reference[1]);
                                        rtxtDataMigration.AppendText(Environment.NewLine + "3. " + reference[2]);
                                        rtxtDataMigration.AppendText(Environment.NewLine);
                                        rtxtDataMigration.AppendText(Environment.NewLine);

                                        for (int i = 0; i < 3; i++)
                                        {
                                            if (result[0, i] == (i + 1).ToString())
                                                rtxtDataMigration.AppendText(Environment.NewLine + (i + 1) + ". " + reference[i] + " - is Missing from '" + oldPc + "'");
                                            else
                                                rtxtDataMigration.AppendText(Environment.NewLine + (i + 1) + ". " + reference[i] + " - Transfer Completed.");
                                        }
                                    }
                                    else if (item == "Quick Access" || item == "Outlook Signatures")
                                    {
                                        bool result = Functions.copyDirectory(newPc, oldPc, username, item);
                                        if (result == true)
                                            rtxtDataMigration.Text = item + " Transfer Done";
                                        else
                                            rtxtDataMigration.Text = "Old PC '" + oldPc + "' doesn't have Signatures";  //Windows has quick access path as default while outlook signatures not
                                    }
                                }
                                else
                                {
                                    string newPcDisplayMessage;
                                    string oldPcDisplayMessage;

                                    if (newPcPing == null)
                                        newPcDisplayMessage = "Offline";
                                    else
                                        newPcDisplayMessage = "Online";

                                    if (oldPcPing == null)
                                        oldPcDisplayMessage = "Offline";
                                    else
                                        oldPcDisplayMessage = "Online";

                                    rtxtDataMigration.Text = "PC Status Error - " + item 
                                                            + "\nNew PC: '" + newPc.ToUpper() + "' Status " + newPcDisplayMessage
                                                            + "\nOld PC: '" + oldPc.ToUpper() + "' Status " + oldPcDisplayMessage;
                                }
                            }
                            else
                            {
                                rtxtDataMigration.Text = "Path is Not existed. Verify Hostnames and Username\n"     //Input wrong user for PC
                                    + "\nNew PC   : " + newPc
                                    + "\nOld PC   : " + oldPc
                                    + "\nNTID     : " + username
                                    + "\nName     : " + usernameAD[0];
                            }
                        }
                        else
                            rtxtDataMigration.Text = item + " Transfer Request Cancelled";
                    }
                }
            }
            catch
            {
                rtxtDataMigration.Text = "Invalid Entry! Please Verify Username and Hostnames\n"
                                        + "\nNew PC: " + "'" + newPc + "'"
                                        + "\nOld PC: " + "'" + oldPc + "'"
                                        + "\nNTID  : " + "'" + username + "'";
            }
        }
    }
}
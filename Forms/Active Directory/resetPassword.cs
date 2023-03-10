using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace desktopDashboard___Y_Lee.Forms
{
    public partial class resetPassword : Form
    {
        public resetPassword()
        {
            InitializeComponent();
        }
        private void btnResetPasswordOK_Click(object sender, EventArgs e)
        {
            string username = txtResetPassword.Text;
            try
            {
                string[] usernameAD = Functions.GetAD(username);
                if (usernameAD == null)
                    rtxtResetPassword.Text = "Invalid Username Entry\n"
                                                + "\nUsername: '" + username.ToUpper() + "' is Incorrect";
                else
                {
                    string answer = MessageBox.Show("Please Confirm Again "
                                                + "\nUser ID: " + usernameAD[2]
                                                + "\nUser name: " + usernameAD[0],
                                                "Desktop Dashboard", MessageBoxButtons.YesNo, MessageBoxIcon.Warning).ToString();

                    if (answer == "Yes")
                    {
                        Functions.resetPassword(username);
                        rtxtResetPassword.Text = "Passowrd Reset Completed\n"
                                            + "\nUser ID      : " + usernameAD[2] 
                                            + "\nUser name    : " + usernameAD[0]
                                            + "\nNew Password : 'Ineos2023'"
                                            + "\n\n User must change password at next logon.";
                    }
                    else
                        rtxtResetPassword.Text = "Password Reset Cancelled";
                }
            }
            catch
            {
                rtxtResetPassword.Text = "Invalid Entry";
            }
        }
    }
}

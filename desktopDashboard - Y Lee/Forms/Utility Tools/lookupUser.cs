using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;


namespace desktopDashboard___Y_Lee.Forms
{
    public partial class lookupUser : Form
    {
        public lookupUser()
        {
            InitializeComponent();
        }

        private void btnLookupUserOK_Click(object sender, EventArgs e)
        {
            string username = txtLookupUser.Text;
            try
            {
                string[] results = Functions.GetAD(username);
                if (results != null)
                {
                    rtxtLookupUser.Text =
                            "Name             : " + results[0]
                        + "\nEmail            : " + results[1]
                        + "\nNTID             : " + results[2]
                        + "\nSite             : " + results[3]
                        + "\nOffice Location  : " + results[4]
                        + "\nTelephone Number : " + results[5]
                        + "\nJob Title        : " + results[6];
                }
                else
                    rtxtLookupUser.Text = "Invalid Entry" + "\n'" + username.ToUpper() + "' is Not Correct";
            }
            catch
            {
                rtxtLookupUser.Text = "Invalid Entry";
            }
        }
    }
}

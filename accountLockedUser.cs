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
    public partial class accountLockedUser : Form
    {
        public accountLockedUser()
        {
            InitializeComponent();
        }

        private void btnAccountLockedUser_Click(object sender, EventArgs e)
        {
            rtxtAccountLockedUser.Clear();
            string site = comboxAccountLockedUser.Text;
            string filter = lbAccountLockedUserTop.Text;

            var (user, ntid, count) = Functions.queryAD(site, filter);
            rtxtAccountLockedUser.AppendText("1. " + user[0]);
            rtxtAccountLockedUser.AppendText(" - " + ntid[0]);

            for (int i = 1; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtAccountLockedUser.AppendText(Environment.NewLine + rtxtCount + ".  " + user[i]);
                rtxtAccountLockedUser.AppendText(" - " + ntid[i]);
            }
            rtxtAccountLockedUser.AppendText(Environment.NewLine);
            rtxtAccountLockedUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}

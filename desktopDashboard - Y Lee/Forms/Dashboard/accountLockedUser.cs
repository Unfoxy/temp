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

            var (name, ntid, email, count) = Functions.queryAD(site, filter);
            rtxtAccountLockedUser.AppendText(string.Format("{0,-4}{1,-26}{2,-41}{3,-20}", "No.", "Name", "Email", "NTID"));

            for (int i = 0; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtAccountLockedUser.AppendText(Environment.NewLine);
                rtxtAccountLockedUser.AppendText(string.Format("{0,-4}{1, -26}{2,-41}{3, -20}", rtxtCount, name[i], email[i], ntid[i]));
            }
            rtxtAccountLockedUser.AppendText(Environment.NewLine);
            rtxtAccountLockedUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}
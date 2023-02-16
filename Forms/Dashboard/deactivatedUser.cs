using desktopDashboard___Y_Lee;
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
    public partial class deactivatedUser : Form
    {
        public deactivatedUser()
        {
            InitializeComponent();
        }
        private void btnDeactivatedUser_Click(object sender, EventArgs e)
        {
            rtxtDeactivatedUser.Clear();
            string site = comboxDeactivatedUser.Text;
            string filter = lbDeactivatedUserTop.Text;

            var (name, ntid, email, count) = Functions.queryAD(site, filter);
            rtxtDeactivatedUser.AppendText(string.Format("{0,-4}{1,-26}{2,-41}{3,-20}", "No.", "Name", "Email", "NTID"));

            for (int i = 0; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtDeactivatedUser.AppendText(Environment.NewLine);
                rtxtDeactivatedUser.AppendText(string.Format("{0,-4}{1, -26}{2,-41}{3, -20}", rtxtCount, name[i], email[i], ntid[i]));
            }
            rtxtDeactivatedUser.AppendText(Environment.NewLine);
            rtxtDeactivatedUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}
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

            var (user, ntid, count) = Functions.queryAD(site, filter);
            rtxtDeactivatedUser.AppendText("1. " + user[0]);
            rtxtDeactivatedUser.AppendText(" - " + ntid[0]);

            for (int i = 1; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtDeactivatedUser.AppendText(Environment.NewLine + rtxtCount + ".  " + user[i]);
                rtxtDeactivatedUser.AppendText(" - " + ntid[i]);
            }
            rtxtDeactivatedUser.AppendText(Environment.NewLine);
            rtxtDeactivatedUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}

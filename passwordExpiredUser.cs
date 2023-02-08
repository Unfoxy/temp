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
    public partial class passwordExpiredUser : Form
    {
        public passwordExpiredUser()
        {
            InitializeComponent();
        }
        private void btnPasswordExpiredUser_Click(object sender, EventArgs e)
        {
            rtxtPasswordExpiredUser.Clear();
            string site = comboxPasswordExpiredUser.Text;
            string filter = lbPasswordExpiredUserTop.Text;

            var (user, ntid, count) = Functions.queryAD(site, filter);
            rtxtPasswordExpiredUser.AppendText("1. " + user[0]);
            rtxtPasswordExpiredUser.AppendText(" - " + ntid[0]);

            for (int i = 1; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtPasswordExpiredUser.AppendText(Environment.NewLine + rtxtCount + ".  " + user[i]);
                rtxtPasswordExpiredUser.AppendText(" - " + ntid[i]);
            }
            rtxtPasswordExpiredUser.AppendText(Environment.NewLine);
            rtxtPasswordExpiredUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}

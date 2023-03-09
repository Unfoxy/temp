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
    public partial class accountExpiredUser : Form
    {
        public accountExpiredUser()
        {
            InitializeComponent();
        }
        private void btnAccountExpiredUser_Click(object sender, EventArgs e)
        {
            rtxtAccountExpiredUser.Clear();
            string site = comboxAccountExpiredUser.Text;
            string filter = lbAccountExpiredUserTop.Text;

            var (name, ntid, email, count) = Functions.queryAD(site, filter);
            rtxtAccountExpiredUser.AppendText(string.Format("{0,-4}{1,-26}{2,-41}{3,-20}", "No.", "Name", "Email", "NTID"));

            for (int i = 0; i < count; i++)
            {
                string rtxtCount = (i + 1).ToString();
                rtxtAccountExpiredUser.AppendText(Environment.NewLine);
                rtxtAccountExpiredUser.AppendText(string.Format("{0,-4}{1, -26}{2,-41}{3, -20}", rtxtCount, name[i],email[i], ntid[i]));
            }
            rtxtAccountExpiredUser.AppendText(Environment.NewLine);
            rtxtAccountExpiredUser.AppendText(Environment.NewLine + "Total Count: " + count);
        }
    }
}